VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_PP097"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File : PS_PP097.cls
'//  Module : 생산관리 > 애드온 > 생산관리
'//  Desc : 분말검사결과확인등록
'//  FormType : PS_PP097
'//  Create Date(Start) : 2017.4.27
'//  Create Date(End) :
'//  Modified Date :
'//  Creator : Hwang Young Su
'//  Company : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form

Public oMat01 As SAPbouiCOM.Matrix
Public oMat02 As SAPbouiCOM.Matrix

Private oDS_PS_PP097L As SAPbouiCOM.DBDataSource '라인(품목분류별규격정보)
Private oDS_PS_PP097M As SAPbouiCOM.DBDataSource '라인(작업지시정보)


Private oLastItemUID01 As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLastColUID01 As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLastColRow01 As Long '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private Type UpdateDatas  '문서에서 삭제된 자료를 저장할 구조체
    InspNo  As Long '삭제한 행의 행번호
    PPYN    As String '삭제한 행의 행번호
    QMYN    As String '삭제한 행의 행번호
End Type

Private UpdateData() As UpdateDatas
Private g_UpdateCount As Integer


'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01 As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_PP097.srf")
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
    
    oFormUniqueID01 = "PS_PP097_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    '//oForm01.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_PP097_CreateItems
    Call PS_PP097_ComboBox_Setting
    Call PS_PP097_Initial_Setting
    Call PS_PP097_CF_ChooseFromList
    Call PS_PP097_EnableMenus
    Call PS_PP097_SetDocument(oFromDocEntry01)
    Call PS_PP097_FormResize
    
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
'        Case et_MATRIX_LINK_PRESSED '//8
'            Call Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, pval, BubbleEvent)
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
            Case "1281": '찾기
            Case "1282": '추가
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
            
            Case "7169": '엑셀 내보내기
            
                '엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
                Call PS_PP097_AddMatrixRow(oMat01.VisualRowCount)
            
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
            Case "1281": '찾기
            Case "1282": '추가
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
            
            Case "7169": '엑셀 내보내기
            
                '엑셀 내보내기 이후 처리
                Call oForm01.Freeze(True)
                Call oDS_PS_PP097L.RemoveRecord(oDS_PS_PP097L.Size - 1)
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
    ElseIf pval.BeforeAction = False Then
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
Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    Dim i As Long
    Dim sQry01 As String
    If pval.BeforeAction = True Then
    
        If pval.ItemUID = "BtnSearch" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If PS_PP097_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                Else
                    
                    '금형이력 정보 초기화
                    oMat02.Clear
                    oDS_PS_PP097M.Clear
                    
                    Call PS_PP097_MTX01 '매트릭스에 데이터 로드
                    
                End If
            End If

        ElseIf pval.ItemUID = "1" Then
        
'            If oForm01.Mode = fm_ADD_MODE Then
'            End If
            
             If g_UpdateCount > 0 Then '행삭제를 한 경우
                            
                For i = 0 To UBound(UpdateData) '구조체 배열에서의 값을 업데이트함.
                
                      sQry01 = "Update [@PS_QM008H] Set U_PPYN= '" & Trim(UpdateData(i).PPYN) & "', U_QMYN=  '" & Trim(UpdateData(i).QMYN) & "' where U_inspNo ='" & Trim(UpdateData(i).InspNo) & "'"
                      Call oRecordSet01.DoQuery(sQry01)
                Next
            
                '삭제행자료 초기화_S
                For i = 0 To UBound(UpdateData)

                    UpdateData(i).InspNo = 0
                    UpdateData(i).PPYN = 0
                    UpdateData(i).QMYN = 0

                Next
                ReDim UpdateData(0)
                '삭제행자료 초기화_E
                
                g_UpdateCount = 0 '삭제행 카운트 초기화
                End If
            

        
        End If
    ElseIf pval.BeforeAction = False Then

    
        If pval.ItemUID = "PS_PP097" Then
        
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
    If pval.BeforeAction = True Then
        
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CardCode", "") '거래처코드 포맷서치 활성
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "ItemCode", "") '품목코드(작번) 포맷서치 활성
        
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
    
        ReDim Preserve UpdateData(g_UpdateCount)
        UpdateData(g_UpdateCount).InspNo = oMat01.Columns("InspNo").Cells(oLastColRow01).Specific.VALUE
        UpdateData(g_UpdateCount).PPYN = oMat01.Columns("PPYN").Cells(oLastColRow01).Specific.VALUE
        UpdateData(g_UpdateCount).QMYN = oMat01.Columns("QMYN").Cells(oLastColRow01).Specific.VALUE
        g_UpdateCount = g_UpdateCount + 1

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
                
                oLastItemUID01 = pval.ItemUID
                oLastColUID01 = pval.ColUID
                oLastColRow01 = pval.Row
            End If
            
        ElseIf pval.ItemUID = "Mat02" Then
            If pval.Row > 0 Then
                
                Call oMat02.SelectRow(pval.Row, True, False)
                
                oLastItemUID01 = pval.ItemUID
                oLastColUID01 = pval.ColUID
                oLastColRow01 = pval.Row
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
    
        If pval.ItemUID = "Mat01" Then '품목대분류별 규격 정보

            If pval.Row = 0 Then

                '정렬
                oMat01.Columns(pval.ColUID).TitleObject.Sortable = True
                oMat01.FlushToDataSource

            Else
                
                Call PS_PP097_MTX02(oMat01.Columns("InspNo").Cells(pval.Row).Specific.VALUE)
                
                
            End If

        End If
    
    ElseIf pval.BeforeAction = False Then
    
    End If
    Exit Sub
Raise_EVENT_DOUBLE_CLICK_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_DOUBLE_CLICK_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_VALIDATE(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_VALIDATE_Error
    Call oForm01.Freeze(True)
    
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    If pval.BeforeAction = True Then
    
    

                
    
    
        If pval.ItemChanged = True Then
        
            If (pval.ItemUID = "CardCode") Then
                oQuery01 = "SELECT CardName, CardCode FROM [OCRD] WHERE CardCode = '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'"
                oRecordSet01.DoQuery oQuery01
                oForm01.Items("CardName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
            ElseIf (pval.ItemUID = "ItemCode") Then
                oQuery01 = "SELECT FrgnName, ItemCode FROM [OITM] WHERE ItemCode = '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'"
                oRecordSet01.DoQuery oQuery01
                oForm01.Items("ItemName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
            End If
            
            oForm01.Items(pval.ItemUID).Click ct_Regular
        End If
    ElseIf pval.BeforeAction = False Then

    End If
    
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
Raise_EVENT_VALIDATE_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_VALIDATE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_MATRIX_LOAD(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_MATRIX_LOAD_Error
    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
        Call PS_PP097_FormItemEnabled
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_PP097_FormResize
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
    ElseIf pval.ItemUID = "Mat02" Then
        If pval.Row > 0 Then
            oLastItemUID01 = pval.ItemUID
            oLastColUID01 = pval.ColUID
            oLastColRow01 = pval.Row
        End If
    ElseIf pval.ItemUID = "Mat03" Then
        If pval.Row > 0 Then
            oLastItemUID01 = pval.ItemUID
            oLastColUID01 = pval.ColUID
            oLastColRow01 = pval.Row
        End If
    ElseIf pval.ItemUID = "Mat04" Then
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
        Set oMat02 = Nothing
    End If
    Exit Sub
Raise_EVENT_FORM_UNLOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_FORM_UNLOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_ROW_DELETE(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_ROW_DELETE_Error
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_PP097_CreateItems() As Boolean
On Error GoTo PS_PP097_CreateItems_Error
    Call oForm01.Freeze(True)

    Set oDS_PS_PP097L = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    Set oDS_PS_PP097M = oForm01.DataSources.DBDataSources("@PS_USERDS02")
    
    '매트릭스 초기화
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    Set oMat02 = oForm01.Items("Mat02").Specific
    oMat02.SelectionMode = ms_NotSupported
    oMat02.AutoResizeColumns
    
    
    ' 사업장_S
    Call oForm01.DataSources.UserDataSources.Add("CardCode", dt_SHORT_TEXT, 100)
    Call oForm01.Items("CardCode").Specific.DataBind.SetBound(True, "", "CardCode")
    '사업장_E
    
    '금형구분_S
    Call oForm01.DataSources.UserDataSources.Add("InspNo", dt_SHORT_TEXT, 100)
    Call oForm01.Items("InspNo").Specific.DataBind.SetBound(True, "", "InspNo")
    '금형구분_E
    
    '구분_S
    Call oForm01.DataSources.UserDataSources.Add("ItemCode", dt_SHORT_TEXT, 100)
    Call oForm01.Items("ItemCode").Specific.DataBind.SetBound(True, "", "ItemCode")
    '구분_E
    
    Call oForm01.DataSources.UserDataSources.Add("FrDate", dt_DATE) '1.조회시작일데이터소스생성
    Call oForm01.Items("FrDate").Specific.DataBind.SetBound(True, "", "FrDate") '2.조회시작일데이터바운드
    
    Call oForm01.DataSources.UserDataSources.Add("ToDate", dt_DATE) '1.조회마지막일데이터소스생성
    Call oForm01.Items("ToDate").Specific.DataBind.SetBound(True, "", "ToDate") '2.조회마지막일데이터바운드
    

    
    
    '실동공수정보 Matrix의 품목코드 필드 Hidden
'    oMat03.Columns("ItemCode").Visible = False
    
    Call oForm01.Freeze(False)
    Exit Function
PS_PP097_CreateItems_Error:
    
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP097_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_PP097_ComboBox_Setting()
On Error GoTo PS_PP097_ComboBox_Setting_Error
    
    Call oForm01.Freeze(True)
        
        
        oMat01.Columns("PPYN").ValidValues.Add "N", "미확인"
        oMat01.Columns("PPYN").ValidValues.Add "Y", "확인"
        
        oMat01.Columns("QMYN").ValidValues.Add "N", "미확인"
        oMat01.Columns("QMYN").ValidValues.Add "Y", "확인"
        

    Call oForm01.Freeze(False)
    Exit Sub
PS_PP097_ComboBox_Setting_Error:
'    Set lRecordSet = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP097_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP097_Initial_Setting()
On Error GoTo PS_PP097_Initial_Setting_Error

    
    g_UpdateCount = 0
    
  '  날짜 설정
    oForm01.Items("ToDate").Specific.VALUE = Format(Date, "YYYYMMDD")
    oForm01.Items("FrDate").Specific.VALUE = Format(Now, "YYYYMM") + "01"
    
    Exit Sub
PS_PP097_Initial_Setting_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP097_Initial_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Sub PS_PP097_CF_ChooseFromList()
On Error GoTo PS_PP097_CF_ChooseFromList_Error
    Exit Sub
PS_PP097_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP097_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP097_FormItemEnabled()
On Error GoTo PS_PP097_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then

    ElseIf (oForm01.Mode = fm_FIND_MODE) Then

    ElseIf (oForm01.Mode = fm_OK_MODE) Then

    End If
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP097_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP097_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP097_AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_PP097_AddMatrixRow_Error
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//행추가여부
        oDS_PS_PP097L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_PP097L.Offset = oRow
'    oDS_PS_PP097L.setValue "U_LineNum", oRow, oRow + 1
    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP097_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP097_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP097_FormClear()
On Error GoTo PS_PP097_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_PP097'", "")
    If DocEntry = 0 Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PS_PP097_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP097_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP097_EnableMenus()
On Error GoTo PS_PP097_EnableMenus_Error
    Exit Sub
PS_PP097_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP097_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP097_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_PP097_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_PP097_FormItemEnabled

    Else

    End If
    Exit Sub
PS_PP097_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP097_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Function PS_PP097_DataValidCheck() As Boolean
On Error GoTo PS_PP097_DataValidCheck_Error
    PS_PP097_DataValidCheck = False
    Dim i As Long


    PS_PP097_DataValidCheck = True
    Exit Function
PS_PP097_DataValidCheck_Error:
    PS_PP097_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PS_PP097_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_PP097_MTX01()
On Error GoTo PS_PP097_MTX01_Error
    '//메트릭스에 데이터 로드
    
    Dim loopCount As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim ItemCode As String  '품목코드
    Dim FrDate As String    '날짜From
    Dim ToDate As String    '날짜To
    Dim InspNo As String    '검사의뢰번호
    Dim CardCode As String  '거래처번호
    
    CardCode = Trim(oForm01.Items("CardCode").Specific.VALUE)
    InspNo = Trim(oForm01.Items("InspNo").Specific.VALUE)
    ItemCode = Trim(oForm01.Items("ItemCode").Specific.VALUE)
    FrDate = Trim(oForm01.Items("FrDate").Specific.VALUE)
    ToDate = Trim(oForm01.Items("ToDate").Specific.VALUE)
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, False) '쿼리를 실행할 때 부터 프로그레스 시작
    
    Call oForm01.Freeze(True)
    
    Query01 = "EXEC PS_PP097_01 '" & CardCode & "','" & ItemCode & "','" & FrDate & "','" & ToDate & "','" & InspNo & "'"
    Call RecordSet01.DoQuery(Query01)
    
    oMat01.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource
    
    If (RecordSet01.RecordCount = 0) Then
        Call oMat01.Clear
        GoTo PS_PP097_MTX01_Exit
    End If
    
    For loopCount = 0 To RecordSet01.RecordCount - 1
        If loopCount <> 0 Then
           Call oDS_PS_PP097L.InsertRecord(loopCount)
        End If
        oDS_PS_PP097L.Offset = loopCount
        Call oDS_PS_PP097L.setValue("U_LineNum", loopCount, loopCount + 1) '라인번호
        Call oDS_PS_PP097L.setValue("U_ColReg01", loopCount, RecordSet01.Fields("CardName").VALUE) '거래처명
        Call oDS_PS_PP097L.setValue("U_ColReg02", loopCount, RecordSet01.Fields("InspNo").VALUE) '검사의뢰번호
        Call oDS_PS_PP097L.setValue("U_ColReg03", loopCount, RecordSet01.Fields("ItemCode").VALUE) '아이템코드
        Call oDS_PS_PP097L.setValue("U_ColReg04", loopCount, RecordSet01.Fields("ItemName").VALUE) '아이템명
        Call oDS_PS_PP097L.setValue("U_ColDt01", loopCount, RecordSet01.Fields("DocDate").VALUE) '검사의뢰날짜
        Call oDS_PS_PP097L.setValue("U_ColQty01", loopCount, RecordSet01.Fields("Weight").VALUE) '검사중량
        Call oDS_PS_PP097L.setValue("U_ColReg05", loopCount, RecordSet01.Fields("PPYN").VALUE) '생산확인
        Call oDS_PS_PP097L.setValue("U_ColReg06", loopCount, RecordSet01.Fields("QMYN").VALUE) '품질확인
     
        RecordSet01.MoveNext
        ProgressBar01.VALUE = ProgressBar01.VALUE + 1
        ProgressBar01.Text = ProgressBar01.VALUE & "/" & RecordSet01.RecordCount & "건 조회중...!"
    Next
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
    
PS_PP097_MTX01_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "W")
    Exit Sub
PS_PP097_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP097_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP097_MTX02(ByVal prmCode As String)
On Error GoTo PS_PP097_MTX02_Error
    '//메트릭스에 데이터 로드
    
    Dim loopCount As Long
    Dim Query01 As String
    Dim Query02 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, False) '쿼리를 실행할 때 부터 프로그레스 시작
    
    Call oForm01.Freeze(True)
                
                
    Query02 = "SELECT U_Sintern as Sintern, U_remark as remark FROM [@PS_QM008H]  WHERE U_inspno='" & prmCode & "'"
    RecordSet01.DoQuery Query02

    oForm01.Items("Sintern").Specific.String = Trim(RecordSet01.Fields("Sintern").VALUE)
    oForm01.Items("remark").Specific.String = Trim(RecordSet01.Fields("remark").VALUE)
    
    
    Query01 = "EXEC PS_PP097_02 '" & prmCode & "'"
    Call RecordSet01.DoQuery(Query01)

    oMat02.Clear
    oMat02.FlushToDataSource
    oMat02.LoadFromDataSource
    
    If (RecordSet01.RecordCount = 0) Then
        Call oMat02.Clear
        GoTo PS_PP097_MTX02_Exit
    End If
    
    For loopCount = 0 To RecordSet01.RecordCount - 1
        If loopCount <> 0 Then
           Call oDS_PS_PP097M.InsertRecord(loopCount)
        End If
        oDS_PS_PP097M.Offset = loopCount
        
        Call oDS_PS_PP097M.setValue("U_LineNum", loopCount, loopCount + 1) '라인번호
        Call oDS_PS_PP097M.setValue("U_ColReg01", loopCount, RecordSet01.Fields("InspItem").VALUE) '이력일자
        Call oDS_PS_PP097M.setValue("U_ColReg02", loopCount, RecordSet01.Fields("InspItNm").VALUE) '완료일자
        Call oDS_PS_PP097M.setValue("U_ColReg03", loopCount, RecordSet01.Fields("InspUnit").VALUE) '두께
        Call oDS_PS_PP097M.setValue("U_ColReg04", loopCount, RecordSet01.Fields("InspSpec").VALUE) '상태
        Call oDS_PS_PP097M.setValue("U_ColReg05", loopCount, RecordSet01.Fields("InspMeth").VALUE) '이력일자
        Call oDS_PS_PP097M.setValue("U_ColQty01", loopCount, RecordSet01.Fields("InspMin").VALUE) '완료일자
        Call oDS_PS_PP097M.setValue("U_ColQty02", loopCount, RecordSet01.Fields("InspMax").VALUE) '두께
        Call oDS_PS_PP097M.setValue("U_ColReg06", loopCount, RecordSet01.Fields("InspBal").VALUE) '상태
        Call oDS_PS_PP097M.setValue("U_ColQty03", loopCount, RecordSet01.Fields("Value").VALUE) '비고

        
        RecordSet01.MoveNext
        ProgressBar01.VALUE = ProgressBar01.VALUE + 1
        ProgressBar01.Text = ProgressBar01.VALUE & "/" & RecordSet01.RecordCount & "건 조회중...!"
    Next
    oMat02.LoadFromDataSource
    oMat02.AutoResizeColumns
    
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
    
PS_PP097_MTX02_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "W")
    Exit Sub
PS_PP097_MTX02_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP097_MTX02_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub



Private Sub PS_PP097_FormResize()
On Error GoTo PS_PP097_FormResize_Error
    
    oForm01.Items("Mat01").Top = 82
    oForm01.Items("Mat01").Height = oForm01.Height / 4

    oForm01.Items("Static01").Top = oForm01.Items("Mat01").Top - 15
    
    oForm01.Items("Mat02").Top = oForm01.Height / 4 + 102
    oForm01.Items("Mat02").Height = oForm01.Height / 4
    oForm01.Items("Static02").Top = oForm01.Items("Mat02").Top - 15
   
    
    oMat01.AutoResizeColumns
    oMat02.AutoResizeColumns
   
    
    Exit Sub
PS_PP097_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP097_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub










