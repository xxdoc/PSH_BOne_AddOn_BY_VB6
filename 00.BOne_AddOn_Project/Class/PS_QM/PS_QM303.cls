VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_QM303"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File : PS_QM303.cls
'//  Module : 품질관리>멀티게이지 품질부적합 관리
'//  Desc : 멀티게이지 품질부적합 해제(2차)
'//  FormType : PS_QM303
'//  Create Date : 2012.01.10
'//  Modified Date :
'//  Creator : Song Myoung gyu
'//  Company : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PS_QM303H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_QM303L As SAPbouiCOM.DBDataSource '등록라인

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
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_QM303.srf")
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
    
    oFormUniqueID01 = "PS_QM303_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
'    oForm01.DataBrowser.BrowseBy = "DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_QM303_CreateItems
    Call PS_QM303_ComboBox_Setting
    Call PS_QM303_Initial_Setting
    Call PS_QM303_CF_ChooseFromList
    Call PS_QM303_EnableMenus
    Call PS_QM303_SetDocument(oFromDocEntry01)
    Call PS_QM303_FormResize
    
    oForm01.EnableMenu ("1283"), False         '// 삭제
    oForm01.EnableMenu ("1287"), False         '// 복제
    oForm01.EnableMenu ("1286"), False        '// 닫기
    oForm01.EnableMenu ("1284"), False        '// 취소
    oForm01.EnableMenu ("1293"), False         '// 행삭제
    
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
            
            Case "7169": '엑셀 내보내기
            
                '엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
                Call PS_QM303_AddMatrixRow(oMat01.VisualRowCount)
            
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": '찾기
                Call PS_QM303_FormItemEnabled '//UDO방식
            Case "1282": '추가
                Call PS_QM303_FormItemEnabled '//UDO방식
                Call PS_QM303_AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                Call PS_QM303_FormItemEnabled
            Case "1287"
            
                oForm01.Freeze True
                Call oDS_PS_QM303H.setValue("DocEntry", 0, "")
                
                For i = 0 To oMat01.VisualRowCount - 1
                    Call oMat01.FlushToDataSource
                    Call oDS_PS_QM303H.setValue("DocEntry", i, "")
                    Call oMat01.LoadFromDataSource
                Next i
                oForm01.Freeze False
                
            Case "7169": '엑셀 내보내기
            
                '엑셀 내보내기 이후 처리
                Call oForm01.Freeze(True)
                Call oDS_PS_QM303L.RemoveRecord(oDS_PS_QM303L.Size - 1)
                Call oMat01.LoadFromDataSource
                Call oForm01.Freeze(False)
                
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
        If pval.ItemUID = "PS_QM303" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        
        If pval.ItemUID = "btnSearch" Then
        
            Call PS_QM303_MTX01
        
        End If
        
        If pval.ItemUID = "btnPrint" Then
        
            Call PS_QM303_Print_Report01

        End If
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_QM303" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
'        If pval.ItemUID = "1" Then
'            If oForm01.Mode = fm_ADD_MODE Then
'                If pval.ActionSuccess = True Then
'                    Call PS_QM303_FormItemEnabled
'                    Call PS_QM303_AddMatrixRow(0, True) '//UDO방식일때
'                End If
'            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
'            ElseIf oForm01.Mode = fm_OK_MODE Then
'                If pval.ActionSuccess = True Then
'                    Call PS_QM303_FormItemEnabled
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
    
        If pval.ItemUID = "Mat01" Then
        
        ElseIf pval.ItemUID = "OCpCode" Then '발생공정

            If pval.CharPressed = 9 Then '탭을 눌렀을 경우만
            
                Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "OCpCode", "")
                
            End If

        ElseIf pval.ItemUID = "OUseMCd" Then '발생공정설비
        
            If pval.CharPressed = 9 Then '탭을 눌렀을 경우만
            
                Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "OUseMCd", "")
                
            End If

        ElseIf pval.ItemUID = "OWkCd1" Then '작업자(주간)
        
            If pval.CharPressed = 9 Then '탭을 눌렀을 경우만
            
                Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "OWkCd1", "")
                
            End If
        
        ElseIf pval.ItemUID = "OWkCd2" Then '작업자(야간)
        
            If pval.CharPressed = 9 Then '탭을 눌렀을 경우만
            
                Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "OWkCd2", "")
                
            End If
        
        ElseIf pval.ItemUID = "RCpCode" Then '원인공정
        
            If pval.CharPressed = 9 Then '탭을 눌렀을 경우만
            
                Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "RCpCode", "")
                
            End If
        
        ElseIf pval.ItemUID = "RUseMCd" Then '원인공정설비
        
            If pval.CharPressed = 9 Then '탭을 눌렀을 경우만
            
                Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "RUseMCd", "")
                
            End If
        
        ElseIf pval.ItemUID = "RWkCd1" Then '작업자(주간)
        
            If pval.CharPressed = 9 Then '탭을 눌렀을 경우만
            
                Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "RWkCd1", "")
                
            End If
        
        ElseIf pval.ItemUID = "RWkCd2" Then '작업자(야간)
        
            If pval.CharPressed = 9 Then '탭을 눌렀을 경우만
                
                Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "RWkCd2", "")
                
            End If
        
        ElseIf pval.ItemUID = "BadCode1" Then '불량분류
        
            If pval.CharPressed = 9 Then '탭을 눌렀을 경우만
            
                Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "BadCode1", "")
                
            End If
        
        ElseIf pval.ItemUID = "BadCode2" Then '불량상세
        
            If pval.CharPressed = 9 Then '탭을 눌렀을 경우만
            
                Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "BadCode2", "")
                
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
            
            If pval.ColUID = "DocEntry" Then
            
                Set oTempClass = New PS_QM300
                
                Call oTempClass.LoadForm(oMat01.Columns("DocEntry").Cells(pval.Row).Specific.VALUE)
            
            End If
            
        End If
    ElseIf pval.BeforeAction = False Then
    
    End If
    
    Set oTempClass = Nothing
    Exit Sub
Raise_EVENT_MATRIX_LINK_PRESSED_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LINK_PRESSED_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_VALIDATE(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_VALIDATE_Error

    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim PS_PH030() As String

    Call oForm01.Freeze(True)
    If pval.BeforeAction = True Then
        If pval.ItemChanged = True Then

            If (pval.ItemUID = "Mat01") Then
            
                oMat01.Columns(pval.ColUID).Cells(pval.Row).Click ct_Regular
                oMat01.AutoResizeColumns
                
            ElseIf (pval.ItemUID = "OCpCode") Then '발생공정
            
                    If oForm01.Items("OCpCode").Specific.VALUE = "CP00000" Then

                        oForm01.Items("OCpName").Specific.VALUE = "원재료불량"

                    Else

                        oForm01.Items("OCpName").Specific.VALUE = MDC_GetData.Get_ReData("U_CpName", "U_CpCode", "[@PS_PP001L]", "'" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'")

                    End If
            
            ElseIf (pval.ItemUID = "OUseMCd") Then '발생공정설비
            
                oForm01.Items("OUseMNm").Specific.VALUE = MDC_GetData.Get_ReData("U_MachName", "U_MachCode", "[@PS_PP130H]", "'" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'")
                
            ElseIf (pval.ItemUID = "OWkCd1") Then '작업자(주간)
            
                oForm01.Items("OWkNm1").Specific.VALUE = MDC_GetData.Get_ReData("U_FullName", "U_MSTCOD", "[OHEM]", "'" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'")
            
            ElseIf (pval.ItemUID = "OWkCd2") Then '작업자(야간)
            
                oForm01.Items("OWkNm2").Specific.VALUE = MDC_GetData.Get_ReData("U_FullName", "U_MSTCOD", "[OHEM]", "'" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'")
            
            ElseIf (pval.ItemUID = "RCpCode") Then '원인공정
            
                If oForm01.Items("RCpCode").Specific.VALUE = "CP00000" Then

                    oForm01.Items("RCpName").Specific.VALUE = "원재료불량"

                Else

                    oForm01.Items("RCpName").Specific.VALUE = MDC_GetData.Get_ReData("U_CpName", "U_CpCode", "[@PS_PP001L]", "'" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'")

                End If
            
            ElseIf (pval.ItemUID = "RUseMCd") Then '원인공정설비
            
                oForm01.Items("RUseMNm").Specific.VALUE = MDC_GetData.Get_ReData("U_MachName", "U_MachCode", "[@PS_PP130H]", "'" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'")
            
            ElseIf (pval.ItemUID = "RWkCd1") Then '작업자(주간)
                
                oForm01.Items("RWkNm1").Specific.VALUE = MDC_GetData.Get_ReData("U_FullName", "U_MSTCOD", "[OHEM]", "'" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'")
            
            ElseIf (pval.ItemUID = "RWkCd2") Then '작업자(야간)
            
                oForm01.Items("RWkNm2").Specific.VALUE = MDC_GetData.Get_ReData("U_FullName", "U_MSTCOD", "[OHEM]", "'" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'")
            
            ElseIf (pval.ItemUID = "BadCode1") Then '불량분류
            
                oForm01.Items("BadName1").Specific.VALUE = MDC_GetData.Get_ReData("U_MidName", "U_MidCode", "[@PS_PP002L]", "'" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'", "AND DocEntry = '1'")
            
            ElseIf (pval.ItemUID = "BadCode2") Then '불량상세
                
                oForm01.Items("BadName2").Specific.VALUE = MDC_GetData.Get_ReData("T1.U_SmalName", "T1.U_SMalCode", "[@PS_PP003H] AS T0 INNER JOIN [@PS_PP003L] AS T1 ON T0.DocEntry = T1.DocEntry", "'" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'", " AND T0.U_MidCode = '" & oForm01.Items("BadCode1").Specific.VALUE & "'")
                
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
        Call PS_QM303_FormItemEnabled
'        Call PS_QM303_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
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
        Call PS_QM303_FormResize
        Call oMat01.AutoResizeColumns
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
            Call oDS_PS_QM303L.RemoveRecord(oDS_PS_QM303L.Size - 1)
            oMat01.LoadFromDataSource
            If oMat01.RowCount = 0 Then
                Call PS_QM303_AddMatrixRow(0)
            Else
                If Trim(oDS_PS_QM303L.GetValue("U_CntcCode", oMat01.RowCount - 1)) <> "" Then
                    Call PS_QM303_AddMatrixRow(oMat01.RowCount)
                End If
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Function PS_QM303_CreateItems() As Boolean
On Error GoTo PS_QM303_CreateItems_Error
   
'    Set oDS_PS_QM303H = oForm01.DataSources.DBDataSources("@PS_QM303H")
    Set oDS_PS_QM303L = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    Set oMat01 = oForm01.Items("Mat01").Specific
    
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    '사업장
    Call oForm01.DataSources.UserDataSources.Add("BPLId", dt_SHORT_TEXT, 100)
    Call oForm01.Items("BPLId").Specific.DataBind.SetBound(True, "", "BPLId")
    
    '발생공정
    Call oForm01.DataSources.UserDataSources.Add("OCpCode", dt_SHORT_TEXT, 100)
    Call oForm01.Items("OCpCode").Specific.DataBind.SetBound(True, "", "OCpCode")
    
    '발생공정명
    Call oForm01.DataSources.UserDataSources.Add("OCpName", dt_SHORT_TEXT, 100)
    Call oForm01.Items("OCpName").Specific.DataBind.SetBound(True, "", "OCpName")
    
    '발생공정설비
    Call oForm01.DataSources.UserDataSources.Add("OUseMCd", dt_SHORT_TEXT, 100)
    Call oForm01.Items("OUseMCd").Specific.DataBind.SetBound(True, "", "OUseMCd")
    
    '발생공정설비명
    Call oForm01.DataSources.UserDataSources.Add("OUseMNm", dt_SHORT_TEXT, 100)
    Call oForm01.Items("OUseMNm").Specific.DataBind.SetBound(True, "", "OUseMNm")
    
    '작업자(주간)
    Call oForm01.DataSources.UserDataSources.Add("OWkCd1", dt_SHORT_TEXT, 100)
    Call oForm01.Items("OWkCd1").Specific.DataBind.SetBound(True, "", "OWkCd1")
    
    '작업자명(주간)
    Call oForm01.DataSources.UserDataSources.Add("OWkNm1", dt_SHORT_TEXT, 100)
    Call oForm01.Items("OWkNm1").Specific.DataBind.SetBound(True, "", "OWkNm1")
    
    '작업자(야간)
    Call oForm01.DataSources.UserDataSources.Add("OWkCd2", dt_SHORT_TEXT, 100)
    Call oForm01.Items("OWkCd2").Specific.DataBind.SetBound(True, "", "OWkCd2")
    
    '작업자명(야간)
    Call oForm01.DataSources.UserDataSources.Add("OWkNm2", dt_SHORT_TEXT, 100)
    Call oForm01.Items("OWkNm2").Specific.DataBind.SetBound(True, "", "OWkNm2")
    
    '원인공정
    Call oForm01.DataSources.UserDataSources.Add("RCpCode", dt_SHORT_TEXT, 100)
    Call oForm01.Items("RCpCode").Specific.DataBind.SetBound(True, "", "RCpCode")
    
    '원인공정명
    Call oForm01.DataSources.UserDataSources.Add("RCpName", dt_SHORT_TEXT, 100)
    Call oForm01.Items("RCpName").Specific.DataBind.SetBound(True, "", "RCpName")
    
    '원인공정설비
    Call oForm01.DataSources.UserDataSources.Add("RUseMCd", dt_SHORT_TEXT, 100)
    Call oForm01.Items("RUseMCd").Specific.DataBind.SetBound(True, "", "RUseMCd")
    
    '원인공정설비명
    Call oForm01.DataSources.UserDataSources.Add("RUseMNm", dt_SHORT_TEXT, 100)
    Call oForm01.Items("RUseMNm").Specific.DataBind.SetBound(True, "", "RUseMNm")
    
    '작업자(주간)
    Call oForm01.DataSources.UserDataSources.Add("RWkCd1", dt_SHORT_TEXT, 100)
    Call oForm01.Items("RWkCd1").Specific.DataBind.SetBound(True, "", "RWkCd1")
    
    '작업자명(주간)
    Call oForm01.DataSources.UserDataSources.Add("RWkNm1", dt_SHORT_TEXT, 100)
    Call oForm01.Items("RWkNm1").Specific.DataBind.SetBound(True, "", "RWkNm1")
    
    '작업자(야간)
    Call oForm01.DataSources.UserDataSources.Add("RWkCd2", dt_SHORT_TEXT, 100)
    Call oForm01.Items("RWkCd2").Specific.DataBind.SetBound(True, "", "RWkCd2")
    
    '작업자명(야간)
    Call oForm01.DataSources.UserDataSources.Add("RWkNm2", dt_SHORT_TEXT, 100)
    Call oForm01.Items("RWkNm2").Specific.DataBind.SetBound(True, "", "RWkNm2")
    
    '불량분류
    Call oForm01.DataSources.UserDataSources.Add("BadCode1", dt_SHORT_TEXT, 100)
    Call oForm01.Items("BadCode1").Specific.DataBind.SetBound(True, "", "BadCode1")
    
    '불량분류명
    Call oForm01.DataSources.UserDataSources.Add("BadName1", dt_SHORT_TEXT, 100)
    Call oForm01.Items("BadName1").Specific.DataBind.SetBound(True, "", "BadName1")
    
    '불량상세
    Call oForm01.DataSources.UserDataSources.Add("BadCode2", dt_SHORT_TEXT, 100)
    Call oForm01.Items("BadCode2").Specific.DataBind.SetBound(True, "", "BadCode2")
    
    '불량상세명
    Call oForm01.DataSources.UserDataSources.Add("BadName2", dt_SHORT_TEXT, 100)
    Call oForm01.Items("BadName2").Specific.DataBind.SetBound(True, "", "BadName2")
    
    '등록일자(Fr)
    Call oForm01.DataSources.UserDataSources.Add("FromDt", dt_DATE)
    Call oForm01.Items("FromDt").Specific.DataBind.SetBound(True, "", "FromDt")
    
    '등록일자(To)
    Call oForm01.DataSources.UserDataSources.Add("ToDt", dt_DATE)
    Call oForm01.Items("ToDt").Specific.DataBind.SetBound(True, "", "ToDt")
    
    '제품명
    Call oForm01.DataSources.UserDataSources.Add("ItemName", dt_SHORT_TEXT, 100)
    Call oForm01.Items("ItemName").Specific.DataBind.SetBound(True, "", "ItemName")
    
    '작지문서번호
    Call oForm01.DataSources.UserDataSources.Add("OrdMgNum", dt_SHORT_TEXT, 100)
    Call oForm01.Items("OrdMgNum").Specific.DataBind.SetBound(True, "", "OrdMgNum")
    
    '온산LotNo
    Call oForm01.DataSources.UserDataSources.Add("BatchNum", dt_SHORT_TEXT, 100)
    Call oForm01.Items("BatchNum").Specific.DataBind.SetBound(True, "", "BatchNum")
    
    '작지Main
    Call oForm01.DataSources.UserDataSources.Add("OrdNum", dt_SHORT_TEXT, 100)
    Call oForm01.Items("OrdNum").Specific.DataBind.SetBound(True, "", "OrdNum")
    
    '해제상태(1차)
    Call oForm01.DataSources.UserDataSources.Add("CclSts1", dt_SHORT_TEXT, 100)
    Call oForm01.Items("CclSts1").Specific.DataBind.SetBound(True, "", "CclSts1")
    
    '해제상태(2차)
    Call oForm01.DataSources.UserDataSources.Add("CclSts2", dt_SHORT_TEXT, 100)
    Call oForm01.Items("CclSts2").Specific.DataBind.SetBound(True, "", "CclSts2")
    
    '문서번호, 라인번호 Hidden 처리
'    oMat01.Columns("DocEntry").Visible = False
'    oMat01.Columns("LineId").Visible = False
    
    Exit Function
PS_QM303_CreateItems_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM303_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_QM303_ComboBox_Setting()
On Error GoTo PS_QM303_ComboBox_Setting_Error
        
    '//콤보에 기본값설정
    Dim oCombo As SAPbouiCOM.ComboBox
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    '사업부
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", False, False)
    
    '해제상태(1차)
    Call oForm01.Items("CclSts1").Specific.ValidValues.Add("%", "전체")
    Call oForm01.Items("CclSts1").Specific.ValidValues.Add("Y", "등록")
    Call oForm01.Items("CclSts1").Specific.ValidValues.Add("N", "해제")
    Call oForm01.Items("CclSts1").Specific.Select(0, psk_Index) '기본으로 등록 세팅
    
    '해제상태(2차)
    Call oForm01.Items("CclSts2").Specific.ValidValues.Add("%", "전체")
    Call oForm01.Items("CclSts2").Specific.ValidValues.Add("Y", "등록")
    Call oForm01.Items("CclSts2").Specific.ValidValues.Add("N", "해제")
    Call oForm01.Items("CclSts2").Specific.Select(0, psk_Index) '기본으로 등록 세팅
    
    '매트릭스 컬럼의 콤보박스 설정_S
    Dim oColumn As SAPbouiCOM.Column

    '4M원인
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_QM303", "Mat01", "4MR", "01", "장비")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_QM303", "Mat01", "4MR", "02", "원소재")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_QM303", "Mat01", "4MR", "03", "작업방법")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_QM303", "Mat01", "4MR", "04", "작업자")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_QM303", "Mat01", "4MR", "05", "작업환경")
    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat01.Columns("4MR"), "PS_QM303", "Mat01", "4MR")
    
    
    Set oColumn = oMat01.Columns("Cancel1")
    Call oColumn.ValidValues.Add("Y", "등록")
    Call oColumn.ValidValues.Add("N", "해제")
    
    Set oColumn = oMat01.Columns("Cancel2")
    Call oColumn.ValidValues.Add("Y", "등록")
    Call oColumn.ValidValues.Add("N", "해제")
    '매트릭스 컬럼의 콤보박스 설정_E

    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oRecordSet01 = Nothing

    Exit Sub
PS_QM303_ComboBox_Setting_Error:
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_QM303_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_QM303_CF_ChooseFromList()
On Error GoTo PS_QM303_CF_ChooseFromList_Error
    Exit Sub
PS_QM303_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM303_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_QM303_FormItemEnabled()
On Error GoTo PS_QM303_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then
        '//각모드에따른 아이템설정
'
        'oForm01.Items("DocEntry").Enabled = False
        oForm01.Items("Mat01").Enabled = True
'        Call PS_QM303_FormClear '//UDO방식
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
PS_QM303_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_QM303_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_QM303_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_QM303_AddMatrixRow_Error
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//행추가여부
        oDS_PS_QM303L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_QM303L.Offset = oRow
    oDS_PS_QM303L.setValue "U_LineNum", oRow, oRow + 1
    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_QM303_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_QM303_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_QM303_FormClear()
On Error GoTo PS_QM303_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_QM303'", "")
    If DocEntry = "" Or DocEntry = "0" Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PS_QM303_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM303_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_QM303_Initial_Setting()
On Error GoTo PS_QM303_Initial_Setting_Error

    ' 사업장
    Call oForm01.Items("BPLId").Specific.Select(MDC_PS_Common.User_BPLId(), psk_ByValue)

    '등록일자
'    oForm01.Items("DocDate").Specific.VALUE = Format(Now, "YYYYMMDD")

    Exit Sub
PS_QM303_Initial_Setting_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM303_Initial_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_QM303_EnableMenus()
On Error GoTo PS_QM303_EnableMenus_Error
    '//메뉴활성화
'    Call oForm01.EnableMenu("1288", True)
'    Call oForm01.EnableMenu("1289", True)
'    Call oForm01.EnableMenu("1290", True)
'    Call oForm01.EnableMenu("1291", True)
    '//Call MDC_GP_EnableMenus(oForm01, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False) '//메뉴설정
    Call MDC_GP_EnableMenus(oForm01, False, False, True, True, False, True, True, True, True, False, False, False, False, False, False) '//메뉴설정
    Exit Sub
PS_QM303_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM303_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_QM303_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_QM303_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_QM303_FormItemEnabled
'        Call PS_QM303_AddMatrixRow(0, True) '//UDO방식일때
    Else
'        oForm01.Mode = fm_FIND_MODE
'        Call PS_QM303_FormItemEnabled
'        oForm01.Items("DocEntry").Specific.VALUE = oFromDocEntry01
'        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_QM303_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM303_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Function PS_QM303_DataValidCheck() As Boolean
On Error GoTo PS_QM303_DataValidCheck_Error
    PS_QM303_DataValidCheck = False
    Dim i As Long
'    If (oForm01.Mode = fm_ADD_MODE) Then
'        Call PS_QM303_FormClear
'    End If
    
    '검사자 미입력시
    If oForm01.Items("ChkWkCd").Specific.VALUE = "" Then
        Call Sbo_Application.SetStatusBarMessage("검사자가 입력되지 않았습니다.", bmt_Short, True)
        PS_QM303_DataValidCheck = False
        Exit Function
    End If
    
    '등록일자 미입력시
    If oForm01.Items("DocDate").Specific.VALUE = "" Then
        Call Sbo_Application.SetStatusBarMessage("등록일자가 입력되지 않았습니다.", bmt_Short, True)
        PS_QM303_DataValidCheck = False
        Exit Function
    End If
    
    '라인정보 미입력 시
    If oMat01.VisualRowCount = 1 Then
        Sbo_Application.SetStatusBarMessage "라인이 존재하지 않습니다.", bmt_Short, True
        PS_QM303_DataValidCheck = False
        Exit Function
    End If
    
    For i = 1 To oMat01.VisualRowCount - 1
        
'        '사번
'        If (oMat01.Columns("CntcCode").Cells(i).Specific.VALUE = "") Then
'            Sbo_Application.SetStatusBarMessage "사번은 필수입니다.", bmt_Short, True
'            oMat01.Columns("CntcCode").Cells(i).Click ct_Regular
'            PS_QM303_DataValidCheck = False
'            Exit Function
'        End If
'
'        '검진구분
'        If (oMat01.Columns("ChkCls").Cells(i).Specific.VALUE = "") Then
'            Sbo_Application.SetStatusBarMessage "검진구분은 필수입니다.", bmt_Short, True
'            oMat01.Columns("ChkCls").Cells(i).Click ct_Regular
'            PS_QM303_DataValidCheck = False
'            Exit Function
'        End If
        
    Next
    
    Call oMat01.FlushToDataSource
    Call oDS_PS_QM303L.RemoveRecord(oDS_PS_QM303L.Size - 1)
    Call oMat01.LoadFromDataSource

'    If (oForm01.Mode = fm_ADD_MODE) Then
'        Call PS_QM303_FormClear
'    End If
    
    PS_QM303_DataValidCheck = True
    Exit Function
PS_QM303_DataValidCheck_Error:
    PS_QM303_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PS_QM303_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Private Sub PS_QM303_MTX01()
On Error GoTo PS_QM303_MTX01_Error

    '//메트릭스에 데이터 로드
    Dim i As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BPLId As String '사업장
    Dim OCpCode As String '발생공정
    Dim OUseMCd As String '발생공정설비
    Dim OWkCd1 As String '작업자(주간)
    Dim OWkCd2 As String '작업자(야간)
    Dim RCpCode As String '원인공정
    Dim RUseMCd As String '원인공정설비
    Dim RWkCd1 As String '작업자(주간)
    Dim RWkCd2 As String '작업자(야간)
    Dim BadCode1 As String '불량분류
    Dim BadCode2 As String '불량상세
    Dim FromDt As String '등록일자(From)
    Dim ToDt As String '등록일자(To)
    Dim ItemName As String '제품명
    Dim OrdMgNum As String '작지문서번호
    Dim BatchNum As String '온산LotNo
    Dim OrdNum As String '작지Main
    Dim CclSts1 As String '해제상태(1차)
    Dim CclSts2 As String '해제상태(2차)
    
    BPLId = oForm01.Items("BPLId").Specific.VALUE '사업장
    OCpCode = oForm01.Items("OCpCode").Specific.VALUE '발생공정
    OUseMCd = oForm01.Items("OUseMCd").Specific.VALUE '발생공정설비
    OWkCd1 = oForm01.Items("OWkCd1").Specific.VALUE '작업자(주간)
    OWkCd2 = oForm01.Items("OWkCd2").Specific.VALUE '작업자(야간)
    RCpCode = oForm01.Items("RCpCode").Specific.VALUE '원인공정
    RUseMCd = oForm01.Items("RUseMCd").Specific.VALUE '원인공정설비
    RWkCd1 = oForm01.Items("RWkCd1").Specific.VALUE '작업자(주간)
    RWkCd2 = oForm01.Items("RWkCd2").Specific.VALUE '작업자(야간)
    BadCode1 = oForm01.Items("BadCode1").Specific.VALUE '불량분류
    BadCode2 = oForm01.Items("BadCode2").Specific.VALUE '불량상세
    FromDt = oForm01.Items("FromDt").Specific.VALUE '등록일자(From)
    ToDt = oForm01.Items("ToDt").Specific.VALUE '등록일자(To)
    ItemName = oForm01.Items("ItemName").Specific.VALUE '제품명
    OrdMgNum = oForm01.Items("OrdMgNum").Specific.VALUE '작지문서번호
    BatchNum = oForm01.Items("BatchNum").Specific.VALUE '온산LotNo
    OrdNum = oForm01.Items("OrdNum").Specific.VALUE '작지Main
    CclSts1 = oForm01.Items("CclSts1").Specific.VALUE '해제상태(1차)
    CclSts2 = oForm01.Items("CclSts2").Specific.VALUE '해제상태(2차)
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, False)
    
    Call oForm01.Freeze(True)
    
    Query01 = "EXEC PS_QM303_01 '" & BPLId & "','" & OCpCode & "','" & OUseMCd & "','" & OWkCd1 & "','" & OWkCd2 & "','" & RCpCode & "','" & RUseMCd & "','" & RWkCd1 & "','" & RWkCd2 & "','" & BadCode1 & "','" & BadCode2 & "','" & FromDt & "','" & ToDt & "','" & ItemName & "','" & OrdMgNum & "','" & BatchNum & "','" & OrdNum & "','" & CclSts1 & "','" & CclSts2 & "'"
    Call RecordSet01.DoQuery(Query01)
    
    oMat01.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource
    
    If (RecordSet01.RecordCount = 0) Then
        Call oMat01.Clear
        GoTo PS_QM303_MTX01_Exit
    End If
    
    For i = 0 To RecordSet01.RecordCount - 1
        If i <> 0 Then
            oDS_PS_QM303L.InsertRecord (i)
        End If
        oDS_PS_QM303L.Offset = i
        Call oDS_PS_QM303L.setValue("U_LineNum", i, i + 1) '라인번호
        Call oDS_PS_QM303L.setValue("U_ColReg01", i, RecordSet01.Fields("ItemCode").VALUE) '제품코드
        Call oDS_PS_QM303L.setValue("U_ColReg02", i, RecordSet01.Fields("ItemName").VALUE) '제품명
        Call oDS_PS_QM303L.setValue("U_ColReg03", i, RecordSet01.Fields("OrdNum").VALUE) '작지Main
        Call oDS_PS_QM303L.setValue("U_ColQty01", i, RecordSet01.Fields("OrdWt").VALUE) '중량
        Call oDS_PS_QM303L.setValue("U_ColQty02", i, RecordSet01.Fields("RealWt").VALUE) '실중량
        Call oDS_PS_QM303L.setValue("U_ColReg06", i, RecordSet01.Fields("4MR").VALUE) '4M원인
        Call oDS_PS_QM303L.setValue("U_ColReg07", i, RecordSet01.Fields("BadCode1").VALUE) '불량분류
        Call oDS_PS_QM303L.setValue("U_ColReg08", i, RecordSet01.Fields("BadName1").VALUE) '불량분류명
        Call oDS_PS_QM303L.setValue("U_ColReg09", i, RecordSet01.Fields("BadCode2").VALUE) '불량상세
        Call oDS_PS_QM303L.setValue("U_ColReg10", i, RecordSet01.Fields("BadName2").VALUE) '불량상세명
        Call oDS_PS_QM303L.setValue("U_ColRgL01", i, RecordSet01.Fields("BadNote").VALUE) '불량사유
        Call oDS_PS_QM303L.setValue("U_ColReg12", i, RecordSet01.Fields("DocDate").VALUE) '등록일자
        Call oDS_PS_QM303L.setValue("U_ColReg13", i, RecordSet01.Fields("Cancel1").VALUE) '1차상태
        Call oDS_PS_QM303L.setValue("U_ColReg14", i, RecordSet01.Fields("Cancel2").VALUE) '2차상태
        
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
PS_QM303_MTX01_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "W")
    Exit Sub
PS_QM303_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_QM303_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub PS_QM303_FormResize()
On Error GoTo PS_QM303_FormResize_Error
    
    Exit Sub
PS_QM303_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM303_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PS_QM303_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PS_QM303_Validate_Error
    PS_QM303_Validate = True
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
PS_QM303_Validate_Exit:
    Set RecordSet01 = Nothing
    Exit Function
PS_QM303_Validate_Error:
    PS_QM303_Validate = False
    Sbo_Application.SetStatusBarMessage "PS_QM303_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_QM303_UPDATE_Cancel()
'******************************************************************************
'Function ID : PS_QM303_UPDATE_Cancel()
'해당모듈 : PS_QM303
'기능 : 해제사유 및 1차 해제 등록
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_QM303_UPDATE_Cancel_Error

    Dim loopCount As Integer
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Dim DocEntry As Integer
    Dim LineId As Integer
    Dim Cancel1 As String
    Dim CancelNt As String
    
    For loopCount = 1 To oMat01.RowCount
    
        DocEntry = oMat01.Columns("DocEntry").Cells(loopCount).Specific.VALUE
        LineId = oMat01.Columns("LineId").Cells(loopCount).Specific.VALUE
        Cancel1 = oMat01.Columns("Cancel1").Cells(loopCount).Specific.VALUE
        CancelNt = oMat01.Columns("CancelNt").Cells(loopCount).Specific.VALUE
    
        Query01 = "EXEC [PS_QM303_02] '" & DocEntry & "','" & LineId & "','" & Cancel1 & "','" & CancelNt & "'"
        
        Call RecordSet01.DoQuery(Query01)
    
    Next
    
    Set RecordSet01 = Nothing

    Exit Sub
PS_QM303_UPDATE_Cancel_Error:
    Set RecordSet01 = Nothing
    Sbo_Application.SetStatusBarMessage "PS_QM303_UPDATE_Cancel_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub



Private Sub PS_QM303_Print_Report01()
On Error GoTo Print_Query_Error
    Dim DocNum As String
    Dim WinTitle                    As String
    Dim ReportName                  As String
    Dim sQry                        As String

    Dim i               As Integer
    Dim ErrNum          As Integer
    Dim Sub_sQry        As String
    
'    Dim BPLId           As String
'    Dim CardCode       As String
'    Dim Pumtxt        As String


    Dim oRecordSet      As SAPbobsCOM.Recordset
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
            
    Call ConnectODBC
    
    '// 조회조건문
    
'    BPLId = Trim(oForm01.Items("BPLId").Specific.VALUE)
'    CardCode = Trim(oForm01.Items("CardCode").Specific.VALUE)
'    Pumtxt = Trim(oForm01.Items("Pumtxt").Specific.VALUE)
'
'    If Pumtxt = "" Then Pumtxt = "%"
'    If CardCode = "" Then CardCode = "%"
'
    
    Dim BPLId As String '사업장
    Dim OCpCode As String '발생공정
    Dim OUseMCd As String '발생공정설비
    Dim OWkCd1 As String '작업자(주간)
    Dim OWkCd2 As String '작업자(야간)
    Dim RCpCode As String '원인공정
    Dim RUseMCd As String '원인공정설비
    Dim RWkCd1 As String '작업자(주간)
    Dim RWkCd2 As String '작업자(야간)
    Dim BadCode1 As String '불량분류
    Dim BadCode2 As String '불량상세
    Dim FromDt As String '등록일자(From)
    Dim ToDt As String '등록일자(To)
    Dim ItemName As String '제품명
    Dim OrdMgNum As String '작지문서번호
    Dim BatchNum As String '온산LotNo
    Dim OrdNum As String '작지Main
    Dim CclSts1 As String '해제상태(1차)
    Dim CclSts2 As String '해제상태(2차)
    
    BPLId = oForm01.Items("BPLId").Specific.VALUE '사업장
    OCpCode = oForm01.Items("OCpCode").Specific.VALUE '발생공정
    OUseMCd = oForm01.Items("OUseMCd").Specific.VALUE '발생공정설비
    OWkCd1 = oForm01.Items("OWkCd1").Specific.VALUE '작업자(주간)
    OWkCd2 = oForm01.Items("OWkCd2").Specific.VALUE '작업자(야간)
    RCpCode = oForm01.Items("RCpCode").Specific.VALUE '원인공정
    RUseMCd = oForm01.Items("RUseMCd").Specific.VALUE '원인공정설비
    RWkCd1 = oForm01.Items("RWkCd1").Specific.VALUE '작업자(주간)
    RWkCd2 = oForm01.Items("RWkCd2").Specific.VALUE '작업자(야간)
    BadCode1 = oForm01.Items("BadCode1").Specific.VALUE '불량분류
    BadCode2 = oForm01.Items("BadCode2").Specific.VALUE '불량상세
    FromDt = oForm01.Items("FromDt").Specific.VALUE '등록일자(From)
    ToDt = oForm01.Items("ToDt").Specific.VALUE '등록일자(To)
    ItemName = oForm01.Items("ItemName").Specific.VALUE '제품명
    OrdMgNum = oForm01.Items("OrdMgNum").Specific.VALUE '작지문서번호
    BatchNum = oForm01.Items("BatchNum").Specific.VALUE '온산LotNo
    OrdNum = oForm01.Items("OrdNum").Specific.VALUE '작지Main
    CclSts1 = oForm01.Items("CclSts1").Specific.VALUE '해제상태(1차)
    CclSts2 = oForm01.Items("CclSts2").Specific.VALUE '해제상태(2차)
    
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    WinTitle = "[PS_QM303] 레포트"
    ReportName = "PS_QM303_01.rpt"
 
    '// Formula 수식필드
    ReDim gRpt_Formula(2)
    ReDim gRpt_Formula_Value(2)
       
    '// SubReport
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)
    
    gRpt_SFormula(1, 1) = ""
    gRpt_SFormula_Value(1, 1) = ""
    
    '/ Procedure 실행"
    sQry = "EXEC PS_QM303_01 '" & BPLId & "','" & OCpCode & "','" & OUseMCd & "','" & OWkCd1 & "','" & OWkCd2 & "','" & RCpCode & "','" & RUseMCd & "','" & RWkCd1 & "','" & RWkCd2 & "','" & BadCode1 & "','" & BadCode2 & "','" & FromDt & "','" & ToDt & "','" & ItemName & "','" & OrdMgNum & "','" & BatchNum & "','" & OrdNum & "','" & CclSts1 & "','" & CclSts2 & "'"

    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
        ErrNum = 1
        GoTo Print_Query_Error
    End If

    '/ Action (sub_query가 있을때는 'Y'로...)/
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") = False Then
    End If
    
    Set oRecordSet = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Print_Query_Error:
    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "Print_Query_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub






