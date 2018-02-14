VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_SD600"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File : PS_SD600.cls
'//  Module : 판매관리>애드온>견적관리
'//  Desc : 견적서등록
'//  FormType : PS_SD600
'//  Create Date : 2013.02.08
'//  Modified Date :
'//  Creator : Song Myoung gyu
'//  Company : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PS_SD600H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_SD600L As SAPbouiCOM.DBDataSource '등록라인

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
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_SD600.srf")
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
    
    oFormUniqueID01 = "PS_SD600_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    oForm01.DataBrowser.BrowseBy = "DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_SD600_CreateItems
    Call PS_SD600_ComboBox_Setting
    Call PS_SD600_Initial_Setting
    Call PS_SD600_CF_ChooseFromList
    Call PS_SD600_EnableMenus
    Call PS_SD600_SetDocument(oFromDocEntry01)
    Call PS_SD600_FormResize
    
    oForm01.EnableMenu ("1283"), False         '// 삭제
    oForm01.EnableMenu ("1287"), False         '// 복제
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
                Call PS_SD600_FormItemEnabled '//UDO방식
            Case "1282": '추가
                Call PS_SD600_FormItemEnabled '//UDO방식
                Call PS_SD600_AddMatrixRow(0, True) '//UDO방식
                Call PS_SD600_Initial_Setting
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                Call PS_SD600_FormItemEnabled
            Case "1287"
            
                oForm01.Freeze True
                Call oDS_PS_SD600H.setValue("DocEntry", 0, "")
                
                For i = 0 To oMat01.VisualRowCount - 1
                    Call oMat01.FlushToDataSource
                    Call oDS_PS_SD600H.setValue("DocEntry", i, "")
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
    
    Dim ErrNum As String '에러번호
    Dim InsertedDocEntry As String '입력된 문서의 문서번호를 리턴

    If pval.BeforeAction = True Then
    
        If pval.ItemUID = "PS_SD600" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        
        If pval.ItemUID = "1" Then
        
            If oForm01.Mode = fm_ADD_MODE Then
            
                If PS_SD600_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                
                Dim QuotNo As String '견적번호
                Dim BPLId As String '사업장
                Dim QuotDt As String '견적일자
                
                BPLId = oForm01.Items("BPLId").Specific.VALUE
                QuotDt = oForm01.Items("QuotDt").Specific.VALUE
                QuotNo = PS_SD600_MakeQutationNo(BPLId, QuotDt)
                
                '견적번호 자동 생성
                If QuotNo = "Overflow" Then
                    ErrNum = "001" '견적번호 오버플로우
                    BubbleEvent = False
                    Exit Sub
                Else
                    Call oDS_PS_SD600H.setValue("U_QuotNo", 0, QuotNo)
                End If
                
                oDocEntry01 = oForm01.Items("DocEntry").Specific.VALUE
                oFormMode01 = oForm01.Mode
                
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            
                If PS_SD600_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                
                oDocEntry01 = oForm01.Items("DocEntry").Specific.VALUE
                oFormMode01 = oForm01.Mode
                
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
            
        ElseIf pval.ItemUID = "btnPrint1" Then
        
            Call PS_SD600_Print_Report01
        
        End If
        
    ElseIf pval.BeforeAction = False Then
    
        If pval.ItemUID = "PS_SD600" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "1" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If pval.ActionSuccess = True Then
                
                '입력 후 재 입력 모드가 아니라 입력한 최종 문서 조회_S
'                    Call PS_SD600_FormItemEnabled
'                    Call PS_SD600_AddMatrixRow(0, True) '//UDO방식일때
'                    Call PS_SD600_Initial_Setting

'                oForm01.Mode = fm_FIND_MODE
'                InsertedDocEntry = PS_SD600_GetMaxDocEntry
'                Call oDS_PS_SD600H.setValue("DocEntry", 0, InsertedDocEntry)
'                oForm01.Items("DocEntry").Enabled = True
'                Call oForm01.Items("1").Click(ct_Regular)
'                oForm01.Items("DocEntry").Enabled = False
                
                oForm01.Mode = fm_OK_MODE
                Call Sbo_Application.ActivateMenuItem("1291") '//이동(최종데이타)
                
                '입력 후 재 입력 모드가 아니라 입력한 최종 문서 조회_E

                End If
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
                If pval.ActionSuccess = True Then
                    Call PS_SD600_FormItemEnabled
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
    
        If pval.ItemUID = "Mat01" Then
        
            If pval.ColUID = "Content1" Then '내용(재질/규격/공정)
            
                If pval.CharPressed = 9 Then '탭을 눌렀을 경우만
            
                    Call MDC_PS_Common.ActiveUserDefineValueAlways(oForm01, pval, BubbleEvent, "Mat01", "Content1") '내용(재질/규격/공정) 포맷서치설정
                
                End If
                
            End If
        
        ElseIf pval.ItemUID = "CntcCode" Then

            Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CntcCode", "") '견적등록자 포맷서치 설정
            
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
        
            If pval.ItemUID = "Mat01" Then
            
                If pval.ColUID = "Class" Then
                
                    Call oMat01.FlushToDataSource
                
                    If oMat01.RowCount = pval.Row And Trim(oDS_PS_SD600L.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
                        Call PS_SD600_AddMatrixRow(pval.Row)
                    End If
                
                End If
            
            End If
            
        End If
        
    End If
    
    Call oMat01.AutoResizeColumns
    
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
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Call oForm01.Freeze(True)
    
    If pval.BeforeAction = True Then
    
        If pval.ItemChanged = True Then

            If (pval.ItemUID = "Mat01") Then
                    
                If pval.ColUID = "Content1" Then '발생공정코드
                
                    oMat01.FlushToDataSource
                    Call oDS_PS_SD600L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                    'Call oDS_PS_SD600L.setValue("U_BadName2", pval.Row - 1, MDC_GetData.Get_ReData("T1.U_SmalName", "T1.U_SMalCode", "[@PS_PP003H] AS T0 INNER JOIN [@PS_PP003L] AS T1 ON T0.DocEntry = T1.DocEntry", "'" & oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE & "'", " AND T0.U_MidCode = '" & oMat01.Columns("BadCode1").Cells(pval.Row).Specific.VALUE & "'"))
                    oMat01.LoadFromDataSource
                    
                ElseIf pval.ColUID = "Content2" Or pval.ColUID = "Price" Then
                
                    oMat01.FlushToDataSource
                    Call oDS_PS_SD600L.setValue("U_Amount", pval.Row - 1, oDS_PS_SD600L.GetValue("U_Content2", pval.Row - 1) * oDS_PS_SD600L.GetValue("U_Price", pval.Row - 1))
                    oMat01.LoadFromDataSource
                    
                End If
                
'                Call oMat01.Columns(pval.ColUID).Cells(pval.Row).Click(ct_Regular)
                oMat01.Columns(pval.ColUID).Cells(pval.Row).Click ct_Regular
                oMat01.AutoResizeColumns
                
            ElseIf (pval.ItemUID = "CntcCode") Then
            
                Call oDS_PS_SD600H.setValue("U_CntcName", 0, MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'"))
                
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
        Call PS_SD600_FormItemEnabled
        Call PS_SD600_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
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
        Call PS_SD600_FormResize
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
            Call oDS_PS_SD600L.RemoveRecord(oDS_PS_SD600L.Size - 1)
            oMat01.LoadFromDataSource
            If oMat01.RowCount = 0 Then
                Call PS_SD600_AddMatrixRow(0)
            Else
                If Trim(oDS_PS_SD600L.GetValue("U_CntcCode", oMat01.RowCount - 1)) <> "" Then
                    Call PS_SD600_AddMatrixRow(oMat01.RowCount)
                End If
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Function PS_SD600_CreateItems() As Boolean
On Error GoTo PS_SD600_CreateItems_Error
   
    Set oDS_PS_SD600H = oForm01.DataSources.DBDataSources("@PS_SD600H")
    Set oDS_PS_SD600L = oForm01.DataSources.DBDataSources("@PS_SD600L")
    Set oMat01 = oForm01.Items("Mat01").Specific
    
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    Exit Function
PS_SD600_CreateItems_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD600_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_SD600_ComboBox_Setting()
On Error GoTo PS_SD600_ComboBox_Setting_Error
        
    '//콤보에 기본값설정
    Dim oCombo As SAPbouiCOM.ComboBox
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_SD600", "Mat01", "UseYN", "Y", "Y")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_SD600", "Mat01", "UseYN", "N", "N")
'    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat01.Columns("UseYN"), "PS_SD600", "Mat01", "UseYN")

    '사업부
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", False, False)
    
    '을병지포함
    Call oForm01.Items("SubYN").Specific.ValidValues.Add("%", "선택")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("SubYN").Specific, "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'M001' ORDER BY Code", "", False, False)
    Call oForm01.Items("SubYN").Specific.Select(0, psk_Index)
    
    '품목구분
    Call oForm01.Items("ItemGrp").Specific.ValidValues.Add("%", "선택")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("ItemGrp").Specific, "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'S002' ORDER BY Code", "", False, False)
    Call oForm01.Items("ItemGrp").Specific.Select(0, psk_Index)
    
    '단위
    Call oForm01.Items("Unit").Specific.ValidValues.Add("%", "선택")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("Unit").Specific, "SELECT Code, Name FROM [@PSH_UOMORG] ORDER BY Code", "", False, False)
    Call oForm01.Items("Unit").Specific.Select(0, psk_Index)
    
    '매트릭스-구분_S
    Dim oColumn As SAPbouiCOM.Column
    
    Set oColumn = oMat01.Columns("Class")
    sQry = "SELECT  U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'S004' ORDER BY U_Minor"
    Call oRecordSet01.DoQuery(sQry)

    Do Until oRecordSet01.EOF
        Call oColumn.ValidValues.Add(oRecordSet01.Fields(0).VALUE, oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    Set oColumn = Nothing
    '매트릭스-구분_E
    
    '매트릭스-단위_S
    Set oColumn = oMat01.Columns("Unit")
    sQry = "SELECT Code, Name FROM [@PSH_UOMORG] ORDER BY Code"
    Call oRecordSet01.DoQuery(sQry)

    Do Until oRecordSet01.EOF
        Call oColumn.ValidValues.Add(oRecordSet01.Fields(0).VALUE, oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    Set oColumn = Nothing
    '매트릭스-단위_E
    
    
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing

    Exit Sub
PS_SD600_ComboBox_Setting_Error:
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_SD600_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_SD600_CF_ChooseFromList()
On Error GoTo PS_SD600_CF_ChooseFromList_Error
    Exit Sub
PS_SD600_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD600_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_SD600_FormItemEnabled()
On Error GoTo PS_SD600_FormItemEnabled_Error

    Call oForm01.Freeze(True)
    
    If (oForm01.Mode = fm_ADD_MODE) Then
    
        '//각모드에따른 아이템설정
        oForm01.Items("DocEntry").Enabled = False
        oForm01.Items("QuotNo").Enabled = False
        oForm01.Items("Mat01").Enabled = True
        Call PS_SD600_FormClear '//UDO방식
        Call oForm01.EnableMenu("1281", True) '//찾기
        Call oForm01.EnableMenu("1282", False) '//추가

    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
    
        '//각모드에따른 아이템설정
        oForm01.Items("DocEntry").Specific.VALUE = ""
        oForm01.Items("DocEntry").Enabled = True
        oForm01.Items("QuotNo").Enabled = True
        oForm01.Items("Mat01").Enabled = False
        Call oForm01.EnableMenu("1281", False) '//찾기
        Call oForm01.EnableMenu("1282", True) '//추가
        
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        '//각모드에따른 아이템설정
        
        oForm01.Items("DocEntry").Enabled = False
        oForm01.Items("QuotNo").Enabled = False
        oForm01.Items("Mat01").Enabled = True
        
    End If
    
    Call oForm01.Freeze(False)
    
    Exit Sub
    
PS_SD600_FormItemEnabled_Error:

    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_SD600_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    
End Sub

Sub PS_SD600_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_SD600_AddMatrixRow_Error
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//행추가여부
        oDS_PS_SD600L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_SD600L.Offset = oRow
    oDS_PS_SD600L.setValue "U_LineNum", oRow, oRow + 1
    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_SD600_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_SD600_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_SD600_FormClear()
On Error GoTo PS_SD600_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_SD600'", "")
    If DocEntry = "" Or DocEntry = "0" Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PS_SD600_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD600_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_SD600_Initial_Setting()
On Error GoTo PS_SD600_Initial_Setting_Error

    Call oForm01.Freeze(True)

    ' 사업장
    Call oForm01.Items("BPLId").Specific.Select(MDC_PS_Common.User_BPLId(), psk_ByValue)

    '견적일자
    oForm01.Items("QuotDt").Specific.VALUE = Format(Now, "YYYYMMDD")
    
    '견적등록자
    oForm01.Items("CntcCode").Specific.VALUE = MDC_PS_Common.User_MSTCOD
    
    '지불조건
    Call oDS_PS_SD600H.setValue("U_PayCon", 0, "귀사 정기 지불")
'    oForm01.Items("PayCon").Specific.VALUE = "귀사 정기 지불"
    
    '납기일
    Call oDS_PS_SD600H.setValue("U_DueDate", 0, "발주 후 60일")
'    oForm01.Items("DueDate").Specific.VALUE = "발주 후 60일"
    
    '물품인도조건
    Call oDS_PS_SD600H.setValue("U_DelCon", 0, "귀사 지정")
'    oForm01.Items("DelCon").Specific.VALUE = "귀사 지정"
    
    '견적유효기간
    Call oDS_PS_SD600H.setValue("U_QuotDue", 0, "30일")
'    oForm01.Items("QuotDue").Specific.VALUE = "30일"
    
    '을병지포함 선택
    Call oForm01.Items("SubYN").Specific.Select("%", psk_ByValue)
    
    '품목구분 선택
    Call oForm01.Items("ItemGrp").Specific.Select("%", psk_ByValue)
    
    '단위 선택
    Call oForm01.Items("Unit").Specific.Select("%", psk_ByValue)
    
    '최초 거래처 텍스트박스 선택
    Call oForm01.Items("CardName").Click(ct_Regular)
    
    Call oForm01.Freeze(False)
    
    Exit Sub
PS_SD600_Initial_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_SD600_Initial_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_SD600_EnableMenus()
On Error GoTo PS_SD600_EnableMenus_Error
    '//메뉴활성화
'    Call oForm01.EnableMenu("1288", True)
'    Call oForm01.EnableMenu("1289", True)
'    Call oForm01.EnableMenu("1290", True)
'    Call oForm01.EnableMenu("1291", True)
    '//Call MDC_GP_EnableMenus(oForm01, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False) '//메뉴설정
    Call MDC_GP_EnableMenus(oForm01, False, False, True, True, False, True, True, True, True, False, False, False, False, False, False) '//메뉴설정
    Exit Sub
PS_SD600_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD600_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_SD600_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_SD600_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_SD600_FormItemEnabled
        Call PS_SD600_AddMatrixRow(0, True) '//UDO방식일때
    Else
        oForm01.Mode = fm_FIND_MODE
        Call PS_SD600_FormItemEnabled
        oForm01.Items("DocEntry").Specific.VALUE = oFromDocEntry01
        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_SD600_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD600_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PS_SD600_DataValidCheck() As Boolean
On Error GoTo PS_SD600_DataValidCheck_Error
    PS_SD600_DataValidCheck = False
    Dim i As Long
    If (oForm01.Mode = fm_ADD_MODE) Then
        Call PS_SD600_FormClear
    End If
    
    '견적번호생성 및 체크 기능
    
    
    '견적일자 미입력시
    If oForm01.Items("QuotDt").Specific.VALUE = "" Then
        Call Sbo_Application.SetStatusBarMessage("견적일자가 입력되지 않았습니다.", bmt_Short, True)
        Call oForm01.Items("QuotDt").Click(ct_Regular)
        PS_SD600_DataValidCheck = False
        Exit Function
    End If
    
    '등록자 미입력시
    If oForm01.Items("CntcCode").Specific.VALUE = "" Then
        Call Sbo_Application.SetStatusBarMessage("등록자가 입력되지 않았습니다.", bmt_Short, True)
        Call oForm01.Items("CntcCode").Click(ct_Regular)
        PS_SD600_DataValidCheck = False
        Exit Function
    End If
    
    '을병지포함
    If oForm01.Items("SubYN").Specific.VALUE = "%" Then
        Call Sbo_Application.SetStatusBarMessage("을병지포함이 선택되지 않았습니다.", bmt_Short, True)
        Call oForm01.Items("SubYN").Click(ct_Regular)
        PS_SD600_DataValidCheck = False
        Exit Function
    End If
    
    '거래처명
    If oForm01.Items("CardName").Specific.VALUE = "" Then
        Call Sbo_Application.SetStatusBarMessage("거래처명이 입력되지 않았습니다.", bmt_Short, True)
        Call oForm01.Items("CardName").Click(ct_Regular)
        PS_SD600_DataValidCheck = False
        Exit Function
    End If
    
    '부서
    If oForm01.Items("CardPart").Specific.VALUE = "" Then
        Call Sbo_Application.SetStatusBarMessage("거래처부서가 입력되지 않았습니다.", bmt_Short, True)
        Call oForm01.Items("CardPart").Click(ct_Regular)
        PS_SD600_DataValidCheck = False
        Exit Function
    End If
    
    '담당자
    If oForm01.Items("CardPrsn").Specific.VALUE = "" Then
        Call Sbo_Application.SetStatusBarMessage("거래처담당자가 입력되지 않았습니다.", bmt_Short, True)
        Call oForm01.Items("CardPrsn").Click(ct_Regular)
        PS_SD600_DataValidCheck = False
        Exit Function
    End If
    
    '견적명
    If oForm01.Items("QuotNm").Specific.VALUE = "" Then
        Call Sbo_Application.SetStatusBarMessage("견적명이 입력되지 않았습니다.", bmt_Short, True)
        Call oForm01.Items("QuotNm").Click(ct_Regular)
        PS_SD600_DataValidCheck = False
        Exit Function
    End If
    
    '품명
    If oForm01.Items("SubYN").Specific.VALUE = "Y" Then '을병지를 포함할 때만 품명 체크
        If oForm01.Items("ItemName").Specific.VALUE = "" Then
            Call Sbo_Application.SetStatusBarMessage("품명이 입력되지 않았습니다.", bmt_Short, True)
            Call oForm01.Items("ItemName").Click(ct_Regular)
            PS_SD600_DataValidCheck = False
            Exit Function
        End If
    End If
    
    '도번/규격
    If oForm01.Items("SubYN").Specific.VALUE = "Y" Then '을병지를 포함할 때만 품명 체크
        If oForm01.Items("Spec").Specific.VALUE = "" Then
            Call Sbo_Application.SetStatusBarMessage("도번/규격이 입력되지 않았습니다.", bmt_Short, True)
            Call oForm01.Items("Spec").Click(ct_Regular)
            PS_SD600_DataValidCheck = False
            Exit Function
        End If
    End If
    
    '품목구분
    If oForm01.Items("ItemGrp").Specific.VALUE = "%" Then
        Call Sbo_Application.SetStatusBarMessage("품목구분이 선택되지 않았습니다.", bmt_Short, True)
        Call oForm01.Items("ItemGrp").Click(ct_Regular)
        PS_SD600_DataValidCheck = False
        Exit Function
    End If
    
    '라인정보 미입력 시
    If oMat01.VisualRowCount = 1 Then
        Sbo_Application.SetStatusBarMessage "라인이 존재하지 않습니다.", bmt_Short, True
        PS_SD600_DataValidCheck = False
        Exit Function
    End If
    
    For i = 1 To oMat01.VisualRowCount - 1
        
        '구분
        If (oMat01.Columns("Class").Cells(i).Specific.VALUE = "") Then
            Sbo_Application.SetStatusBarMessage "구분은 필수입니다.", bmt_Short, True
            oMat01.Columns("Class").Cells(i).Click ct_Regular
            PS_SD600_DataValidCheck = False
            Exit Function
        End If

        '내용(재질/규격/공정)
        If (oMat01.Columns("Content1").Cells(i).Specific.VALUE = "") Then
            Sbo_Application.SetStatusBarMessage "내용(재질/규격/공정)은 필수입니다.", bmt_Short, True
            oMat01.Columns("Content1").Cells(i).Click ct_Regular
            PS_SD600_DataValidCheck = False
            Exit Function
        End If
        
        '수량(중량/공수)
        If (oMat01.Columns("Content2").Cells(i).Specific.VALUE = "") Then
            Sbo_Application.SetStatusBarMessage "수량(중량/공수)은 필수입니다.", bmt_Short, True
            oMat01.Columns("Content2").Cells(i).Click ct_Regular
            PS_SD600_DataValidCheck = False
            Exit Function
        End If
        
        '단위
        If (oMat01.Columns("Unit").Cells(i).Specific.VALUE = "") Then
            Sbo_Application.SetStatusBarMessage "검진구분은 필수입니다.", bmt_Short, True
            oMat01.Columns("Unit").Cells(i).Click ct_Regular
            PS_SD600_DataValidCheck = False
            Exit Function
        End If
        
        '단가
        If (oMat01.Columns("Price").Cells(i).Specific.VALUE = "") Then
            Sbo_Application.SetStatusBarMessage "단가는 필수입니다.", bmt_Short, True
            oMat01.Columns("Price").Cells(i).Click ct_Regular
            PS_SD600_DataValidCheck = False
            Exit Function
        End If
        
        '금액
        If (oMat01.Columns("Amount").Cells(i).Specific.VALUE = "") Then
            Sbo_Application.SetStatusBarMessage "금액은 필수입니다.", bmt_Short, True
            oMat01.Columns("Amount").Cells(i).Click ct_Regular
            PS_SD600_DataValidCheck = False
            Exit Function
        End If

    Next
    
    Call oMat01.FlushToDataSource
    Call oDS_PS_SD600L.RemoveRecord(oDS_PS_SD600L.Size - 1)
    Call oMat01.LoadFromDataSource

    If (oForm01.Mode = fm_ADD_MODE) Then
        Call PS_SD600_FormClear
    End If
    
    PS_SD600_DataValidCheck = True
    Exit Function
PS_SD600_DataValidCheck_Error:
    PS_SD600_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PS_SD600_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_SD600_MTX01()
On Error GoTo PS_SD600_MTX01_Error
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
        GoTo PS_SD600_MTX01_Exit
    End If
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, False)
    
    For i = 0 To RecordSet01.RecordCount - 1
        If i <> 0 Then
            oDS_PS_SD600L.InsertRecord (i)
        End If
        oDS_PS_SD600L.Offset = i
        Call oDS_PS_SD600L.setValue("U_COL01", i, RecordSet01.Fields(0).VALUE)
        Call oDS_PS_SD600L.setValue("U_COL02", i, RecordSet01.Fields(1).VALUE)
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
PS_SD600_MTX01_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Exit Sub
PS_SD600_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_SD600_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub PS_SD600_FormResize()
On Error GoTo PS_SD600_FormResize_Error
    
    Exit Sub
PS_SD600_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_SD600_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PS_SD600_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PS_SD600_Validate_Error
    PS_SD600_Validate = True
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
PS_SD600_Validate_Exit:
    Set RecordSet01 = Nothing
    Exit Function
PS_SD600_Validate_Error:
    PS_SD600_Validate = False
    Sbo_Application.SetStatusBarMessage "PS_SD600_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Function PS_SD600_MakeQutationNo(ByVal prmBPLId As String, ByVal prmQuotDt As String) As String
'******************************************************************************
'Function ID : PS_SD600_MakeQutationNo()
'해당모듈 : PS_SD600
'기능 : 견적번호 자동생성
'인수 : prmBPLID(사업장), prmQuotDt(견적일자)
'반환값 : 견적번호
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_SD600_MakeQutationNo_Error

    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    sQry = "EXEC PS_SD600_90 '" & prmBPLId & "','" & prmQuotDt & "'"
    
    Call oRecordSet01.DoQuery(sQry)
    
    PS_SD600_MakeQutationNo = oRecordSet01.Fields("QuotNo").VALUE

    Set oRecordSet01 = Nothing
    
    Exit Function

PS_SD600_MakeQutationNo_Error:

    Set oRecordSet01 = Nothing
    Sbo_Application.SetStatusBarMessage "PS_SD600_MakeQutationNo_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True

End Function

Private Function PS_SD600_GetMaxDocEntry() As String
'******************************************************************************
'Function ID : PS_SD600_GetMaxDocEntry()
'해당모듈 : PS_SD600
'기능 : 가장 최근 등록된 문서번호를 조회
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_SD600_GetMaxDocEntry_Error

    Dim loopCount As Integer
    Dim sQry As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)

    sQry = "SELECT MAX(AutoKey) - 1 AS [MaxDocEntry] FROM ONNM WHERE ObjectCode = 'PS_SD600'"
    
    Call oRecordSet.DoQuery(sQry)
    
    PS_SD600_GetMaxDocEntry = oRecordSet.Fields("MaxDocEntry").VALUE
    
    Set oRecordSet = Nothing
    Exit Function
PS_SD600_GetMaxDocEntry_Error:
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "PS_SD600_GetMaxDocEntry_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_SD600_Print_Report01()
'******************************************************************************
'Function ID : PS_SD600_GetMaxDocEntry()
'해당모듈 : PS_SD600
'기능 : 견적서 갑지 리포트 출력
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_SD600_Print_Report01_Error

    Dim DocNum As String
    Dim WinTitle As String
    Dim ReportName As String
    Dim sQry As String

    Dim i As Integer
    Dim ErrNum As Integer
    Dim Sub_sQry As String

    Dim oRecordSet As SAPbobsCOM.Recordset

    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)

    Call ConnectODBC

    '// 조회조건문
    Dim QuotNo As String '견적번호
    Dim SubYN As String '을병지포함

    QuotNo = Trim(oForm01.Items("QuotNo").Specific.VALUE)
    SubYN = Trim(oForm01.Items("SubYN").Specific.VALUE)

    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    WinTitle = "견적서 갑지"
    ReportName = "PS_SD600_51.rpt"

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
    sQry = "EXEC PS_SD600_01 '" & QuotNo & "','" & SubYN & "'"

    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
        ErrNum = 1
        GoTo PS_SD600_Print_Report01_Error
    End If

    '/ Action (sub_query가 있을때는 'Y'로...)/
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") = False Then
    End If

    Set oRecordSet = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_SD600_Print_Report01_Error:
    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "PS_SD600_Print_Report01_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub
