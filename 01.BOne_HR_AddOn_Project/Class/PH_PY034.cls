VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY034"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File               : PH_PY034.cls
'//  Module             : 인사관리>기타관리
'//  Desc               : 공용분개처리
'//  FormType           : PH_PY034
'//  Create Date(Start) : 2013.10.15
'//  Create Date(End)   :
'//  Creator            : Song Myoung gyu
'//  Modified Date      :
'//  Modifier           :
'//  Company            : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm           As SAPbouiCOM.Form

Public oMat01          As SAPbouiCOM.Matrix
Private oDS_PH_PY034A  As SAPbouiCOM.DBDataSource '헤더
Private oDS_PH_PY034B  As SAPbouiCOM.DBDataSource '라인

Private oLast_Item_UID As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLast_Col_UID  As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLast_Col_Row  As Long '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private oSeq&
'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error

    Dim i         As Long
    Dim oInnerXml As String
    Dim oXmlDoc   As New MSXML2.DOMDocument
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY034.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    
    '매트릭스의 타이틀높이와 셀높이를 고정
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    
    oFormUniqueID01 = "PH_PY034_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01, "PH_PY034" '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    '폼 할당
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID01)
    
    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
    oForm.DataBrowser.BrowseBy = "DocEntry" '//UDO방식일때
        
    oForm.Freeze True
    Call PH_PY034_CreateItems
    Call PH_PY034_ComboBox_Setting
    Call PH_PY034_CF_ChooseFromList
    Call PH_PY034_EnableMenus
    Call PH_PY034_SetDocument(oFromDocEntry01)
    Call PH_PY034_FormResize
    Call PH_PY034_FormItemEnabled

    oForm.Update
    oForm.Freeze False
    
    oForm.Visible = True
    Set oXmlDoc = Nothing
    
    Exit Sub
LoadForm_Error:
    oForm.Update
    oForm.Freeze False
    Set oXmlDoc = Nothing
    Set oForm = Nothing
    Sbo_Application.SetStatusBarMessage "Form_Load Error:" & Err.Description, bmt_Short, True
End Sub

'/아이템 변경 이벤트
Public Sub Raise_FormItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_FormItemEvent_Error

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
Raise_FormItemEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_FormItemEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

'
'Public Sub Raise_ItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
'On Error GoTo Raise_ItemEvent_Error
'Dim i As Integer
'Dim sQry                As String
'Dim oRecordSet01        As SAPbobsCOM.Recordset
'Dim TempForm01 As Variant
'Dim ErrNum As Integer
'
'Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)   '// 객체 정의 및 데이터 할당
'
'    If (pval.BeforeAction = True) Then '//BeforeAction = True
'        Select Case pval.EventType
'
'            Case et_ITEM_PRESSED: '//1
'            If pval.ItemUID = "1" Then
'                If oForm.Mode = fm_ADD_MODE Or oForm.Mode = fm_UPDATE_MODE Then
'                   If HeaderSpaceLineDel = False Then
'                        BubbleEvent = False                   ' BubbleEvent = True 이면, 사용자에게 제어권을 넘겨준다. BeforeAction = True일 경우만 쓴다.
'                        Exit Sub
'                    End If
'                    If MatrixSpaceLineDel = False Then
'                        BubbleEvent = False
'                        Exit Sub
'                    End If
'                End If
'            ElseIf pval.ItemUID = "Btn01" Then      '// 상각자료 불러오기
'                Call MTX01
'            ElseIf pval.ItemUID = "Btn02" Then      '// DI API - 분개 생성
'               If oForm.Mode = fm_OK_MODE Then
'                  If oForm.Items("JdtDate").Specific.VALUE = "" Then
'                     ErrNum = 1
'                     Call Item_Error_Message(ErrNum)
'                     BubbleEvent = False
'                     Exit Sub
'                  ElseIf oForm.Items("Status").Specific.VALUE = "C" Then
'                     ErrNum = 2
'                     Call Item_Error_Message(ErrNum)
'                     BubbleEvent = False
'                     Exit Sub
'                  Else
'                    If Create_oJournalEntries(1) = False Then
'                      BubbleEvent = False
'                      Exit Sub
'                    End If
'                  End If
'
'               Else
'                   Call MDC_Com.MDC_GF_Message("먼저 저장한 후 분개 처리 바랍니다.", "W")
'                   BubbleEvent = False
'                   Exit Sub
'               End If
'
'            ElseIf pval.ItemUID = "Btn03" Then      '// DI API - 분개 취소
'               If oForm.Mode = fm_OK_MODE Then
'                  If oForm.Items("JdtDate").Specific.VALUE = "" Then
'                     ErrNum = 1
'                     Call Item_Error_Message(ErrNum)
'                     BubbleEvent = False
'                     Exit Sub
'                  ElseIf oForm.Items("JdtCC").Specific.VALUE <> "Y" Then
'                     ErrNum = 3
'                     Call Item_Error_Message(ErrNum)
'                     BubbleEvent = False
'                     Exit Sub
'                  ElseIf oForm.Items("Status").Specific.VALUE = "C" Then
'                     ErrNum = 2
'                     Call Item_Error_Message(ErrNum)
'                     BubbleEvent = False
'                     Exit Sub
'                  Else
'                     If Cancel_oJournalEntries(1) = False Then
'                        BubbleEvent = False
'                        Exit Sub
'                     End If
'                  End If
'               Else
'                     Call MDC_Com.MDC_GF_Message("먼저 저장한 후 분개 처리 바랍니다.", "W")
'                     BubbleEvent = False
'                     Exit Sub
'               End If
'
'            Else
'                If pval.ItemChanged = True Then
'
'                End If
'            End If
'
'
'            Case et_KEY_DOWN: '//2
'                ' 거래처코드
'
'
'            Case et_COMBO_SELECT: '//5
'            Case et_CLICK: '//6
'            Case et_DOUBLE_CLICK: '//7
'            If pval.ItemChanged = True Then
'            End If
'
'            Case et_MATRIX_LINK_PRESSED '//8
'            Case et_VALIDATE: '//10
'
'
'
'            Case et_MATRIX_LOAD: '//11
'            Case et_FORM_ACTIVATE: '//18
'            Case et_FORM_DEACTIVATE: '//19
'            Case et_FORM_RESIZE '//20
'            Case et_CHOOSE_FROM_LIST '//27
'            Case et_GOT_FOCUS: '//3
'                oLast_Item_UID = pval.ItemUID
'            Case et_LOST_FOCUS: '//4
'            Case et_FORM_UNLOAD: '//17
'        End Select
'
'    ElseIf (pval.BeforeAction = False) Then         '//BeforeAction = False
'        Select Case pval.EventType
'            Case et_ITEM_PRESSED: '//1
'
'             ' 저장 후 추가 가능처리
'                If pval.ItemUID = "1" Then
'                    If oForm.Mode = fm_ADD_MODE And pval.Action_Success = True Then
'                          oForm.Mode = fm_OK_MODE
'                          Call Sbo_Application.ActivateMenuItem("1282")
'                    ElseIf oForm.Mode = fm_ADD_MODE And pval.Action_Success = False Then
'                        FormItemEnabled
'                        AddMatrixRow 1, oMat01.RowCount, True
'                    End If
'                End If
'            Case et_KEY_DOWN: '//2
'                If pval.Action_Success = True Then
'                    oSeq = 1
'                End If
'            Case et_COMBO_SELECT: '//5
'
'
'            Case et_CLICK: '//6
'            Case et_DOUBLE_CLICK: '//7
'            Case et_MATRIX_LINK_PRESSED '//8
'            Case et_VALIDATE: '//10
'
'
'            Case et_MATRIX_LOAD: '//11
''                AddMatrixRow 1, oMat01.VisualRowCount, True
'            Case et_FORM_ACTIVATE: '//18
'            If oSeq = 1 Then
'                oSeq = 0
'            End If
'            Case et_FORM_DEACTIVATE: '//19
'            Case et_FORM_RESIZE '//20
'            Case et_CHOOSE_FROM_LIST '//27
'
'            Case et_GOT_FOCUS: '//3
'                oLast_Item_UID = pval.ItemUID
'
'            Case et_LOST_FOCUS: '//4
'            Case et_FORM_UNLOAD: '//17
'                RemoveForms oFormUniqueID01
'                Set oForm = Nothing
'                Set oMat01 = Nothing
'        End Select
'      End If
'
'    Exit Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Raise_ItemEvent_Error:
'        Sbo_Application.SetStatusBarMessage "Raise_ItemEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
'End Sub
'
'Public Sub Raise_MenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
'On Error GoTo Raise_MenuEvent_Error
'  Dim i&
'    If (pval.BeforeAction = True) Then '//BeforeAction = True
'        Select Case pval.MenuUID
'            Case "1284": '취소
'            Case "1286": '닫기
'            Case "1293": '행닫기
'
'            Case "1281": '찾기
'            Case "1282": '추가
'            Case "1288", "1289", "1290", "1291": '레코드이동버튼
'
'        End Select
'    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
'        Select Case pval.MenuUID
'            Case "1284": '취소
'            Case "1286": '닫기
'            Case "1281": '찾기
'                FormItemEnabled
''                oForm.Items("ItemCode").Click ct_Regular
'            Case "1282": '추가
'                FormItemEnabled
'                FormClear
'                AddMatrixRow 0, oMat01.RowCount, True
'                oForm.Items("YM").CLICK ct_Collapsed
'
'            Case "1288", "1289", "1290", "1291": '레코드이동버튼
'                FormItemEnabled
'                If oMat01.VisualRowCount > 0 Then
'                    If oMat01.Columns("AcctCode").Cells(oMat01.VisualRowCount).Specific.VALUE <> "" Then
'                        AddMatrixRow 1, oMat01.RowCount, True
'                    End If
'                End If
'            Case "1293": '행닫기
'
'        End Select
'    End If
'    Exit Sub
'Raise_MenuEvent_Error:
'    Sbo_Application.SetStatusBarMessage "Raise_MenuEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
'End Sub
'
'Public Sub Raise_FormDataEvent(ByRef FormUID As String, ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
'On Error GoTo Raise_FormDataEvent_Error
'    If (BusinessObjectInfo.BeforeAction = True) Then '//BeforeAction = True
'        Select Case BusinessObjectInfo.EventType
'            Case et_FORM_DATA_LOAD: '//33
'            Case et_FORM_DATA_ADD:  '//34 - 추가
''            FormClear
''            If Create_oJournalEntries(2) = False Then
''                BubbleEvent = False
''                Exit Sub
''            End If
'            Case et_FORM_DATA_UPDATE: '//35 - 업데이트
'            Case et_FORM_DATA_DELETE: '//36
'
'        End Select
'    ElseIf (BusinessObjectInfo.BeforeAction = False) Then '//BeforeAction = False
'        Select Case BusinessObjectInfo.EventType
'            Case et_FORM_DATA_LOAD: '//33
'            Case et_FORM_DATA_ADD: '//34
'            Case et_FORM_DATA_UPDATE: '//35
'            Case et_FORM_DATA_DELETE: '//36
'        End Select
'    End If
'    Exit Sub
'Raise_FormDataEvent_Error:
'    Sbo_Application.SetStatusBarMessage "Raise_FormDataEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
'End Sub
'
'Public Sub Raise_RightClickEvent(ByRef FormUID As String, ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
'On Error GoTo Raise_RightClickEvent_Error
'    If (eventInfo.BeforeAction = True) Then
'        '//작업
'    ElseIf (eventInfo.BeforeAction = False) Then
'        '//작업
'    End If
'    Exit Sub
'Raise_RightClickEvent_Error:
'    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
'End Sub
'
'Function CreateItems() As Boolean
'On Error GoTo CreateItems_Error
'    Dim sQry As String
'    Dim oRecordSet01 As SAPbobsCOM.Recordset
'
'    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
'
'    Set oDS_PH_PY034H = oForm.DataSources.DBDataSources("@PH_PY034H")
'    Set oDS_PH_PY034L = oForm.DataSources.DBDataSources("@PH_PY034L")
'    Set oMat01 = oForm.Items("Mat01").Specific                        '// 매트릭스 데이터 셋
'
'     oDS_PH_PY034H.setValue "U_YM", 0, Format(Now, "yyyymm")
'     oDS_PH_PY034H.setValue "U_JdtDate", 0, Format(Now, "yyyymmdd")
'
'    '// 사업장 리스트
'    sQry = "SELECT BPLId, BPLName FROM [OBPL]"
'
'    oRecordSet01.DoQuery sQry
'
'    Do Until oRecordSet01.EOF
'        oForm.Items("BPLId").Specific.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
'        oRecordSet01.MoveNext
'    Loop
'
'
'
'    Set oRecordSet01 = Nothing
'    Exit Function
'
'CreateItems_Error:
'    Set oRecordSet01 = Nothing
'    Sbo_Application.SetStatusBarMessage "CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
'End Function

Public Sub Raise_FormMenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_FormMenuEvent_Error

    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = oCompany.GetBusinessObject(BoRecordset)

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
                Call PH_PY034_FormItemEnabled
                Call PH_PY034_AddMatrixRow
            Case "1282": '추가
                Call PH_PY034_FormItemEnabled
                Call PH_PY034_AddMatrixRow
            
'                oMat01.Clear
'                oDS_PH_PY032A.Clear

'                Call PH_PY032_LoadCaption
'                Call PH_PY032_FormItemEnabled
                '//Call PH_PY032_FormItemEnabled '//UDO방식
                '//Call PH_PY032_AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                '//Call PH_PY032_FormItemEnabled
        End Select
    End If
    Exit Sub
Raise_FormMenuEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_FormMenuEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
            oLast_Item_UID = pval.ItemUID
            oLast_Col_UID = pval.ColUID
            oLast_Col_Row = pval.Row
        End If
    Else
        oLast_Item_UID = pval.ItemUID
        oLast_Col_UID = ""
        oLast_Col_Row = 0
    End If
    Exit Sub
Raise_RightClickEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_ITEM_PRESSED(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_ITEM_PRESSED_Error

    Dim ErrNum As Integer

    If pval.BeforeAction = True Then
    
        If pval.ItemUID = "1" Then
        
            If oForm.Mode = fm_ADD_MODE Then
            
                If PH_PY034_DataValidCheck = False Then
                    BubbleEvent = False
                End If
                
            ElseIf oForm.Mode = fm_UPDATE_MODE Then
            
                If PH_PY034_DataValidCheck = False Then
                    BubbleEvent = False
                End If
                
            ElseIf oForm.Mode = fm_OK_MODE Then
            
            End If
        
        ElseIf pval.ItemUID = "Btn01" Then '조회버튼
        
            Call PH_PY034_MTX01
        
        ElseIf pval.ItemUID = "Btn02" Then '분개확정 버튼
        
            If oForm.Mode = fm_OK_MODE Then
            
                If oForm.Items("JdtDate").Specific.VALUE = "" Then
                    ErrNum = 1
                    Call PH_PY034_Item_Error_Message(ErrNum)
                    BubbleEvent = False
                    Exit Sub
                ElseIf oForm.Items("Status").Specific.VALUE = "C" Then
                    ErrNum = 2
                    Call PH_PY034_Item_Error_Message(ErrNum)
                    BubbleEvent = False
                    Exit Sub
                Else
                    If PH_PY034_Create_oJournalEntries(1) = False Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                End If

            Else
                Call MDC_Com.MDC_GF_Message("먼저 갱신한 후 분개 처리 바랍니다.", "W")
                BubbleEvent = False
                Exit Sub
            End If

        ElseIf pval.ItemUID = "Btn03" Then '분개취소 버튼
        
            If oForm.Mode = fm_OK_MODE Then
                If oForm.Items("JdtDate").Specific.VALUE = "" Then
                    ErrNum = 1
                    Call PH_PY034_Item_Error_Message(ErrNum)
                    BubbleEvent = False
                    Exit Sub
                ElseIf oForm.Items("JdtCC").Specific.VALUE <> "Y" Then
                    ErrNum = 3
                    Call PH_PY034_Item_Error_Message(ErrNum)
                    BubbleEvent = False
                    Exit Sub
                ElseIf oForm.Items("Status").Specific.VALUE = "C" Then
                    ErrNum = 2
                    Call PH_PY034_Item_Error_Message(ErrNum)
                    BubbleEvent = False
                    Exit Sub
               Else
                    If PH_PY034_Cancel_oJournalEntries(1) = False Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                End If
            Else
                Call MDC_Com.MDC_GF_Message("먼저 갱신한 후 분개 처리 바랍니다.", "W")
                BubbleEvent = False
                Exit Sub
            End If
            
        End If
        
    ElseIf pval.BeforeAction = False Then
    
        If pval.ItemUID = "1" Then
        
            If oForm.Mode = fm_ADD_MODE Then
            
                If pval.ActionSuccess = True Then
                    oForm.Mode = fm_OK_MODE
                    Call Sbo_Application.ActivateMenuItem("1291") '//이동(최종데이타)
                ElseIf pval.Action_Success = False Then
                    Call PH_PY034_FormItemEnabled
                    Call PH_PY034_AddMatrixRow
                End If
                
            ElseIf oForm.Mode = fm_UPDATE_MODE Then
            
                If pval.ActionSuccess = True Then
                    Call PH_PY034_FormItemEnabled
                    Call PH_PY034_AddMatrixRow
                End If
                
            ElseIf oForm.Mode = fm_OK_MODE Then
            
                If pval.ActionSuccess = True Then
                    Call PH_PY034_FormItemEnabled
                End If
                
            End If
            
        End If
        
    End If


'    If pval.BeforeAction = True Then
'
'        If pval.ItemUID = "PH_PY034" Then
'            If oForm.Mode = fm_ADD_MODE Then
'            ElseIf oForm.Mode = fm_UPDATE_MODE Then
'            ElseIf oForm.Mode = fm_OK_MODE Then
'            End If
'        End If
'
'        If pval.ItemUID = "1" Then '/추가/확인 버튼클릭
'
'            If oForm.Mode = fm_ADD_MODE Then
'
'                If PH_PY034_HeaderSpaceLineDel() = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If
'
''                If PH_PY034_DataCheck() = False Then
''                    BubbleEvent = False
''                    Exit Sub
''                End If
'
''                Call PH_PY032_FormReset
'                oForm.Mode = fm_ADD_MODE
'
''                Call PH_PY032_LoadCaption
''                Call PH_PY032_MTX01
'
'                oLast_Mode = oForm.Mode
'
'            ElseIf oForm.Mode = fm_UPDATE_MODE Then
'
'                If PH_PY034_HeaderSpaceLineDel = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If
'
''                If PH_PY032_DataCheck() = False Then
''                    BubbleEvent = False
''                    Exit Sub
''                End If
'
'                If PH_PY034_UpdateData() = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If
'
'                Call PH_PY034_FormReset
'                oForm.Mode = fm_ADD_MODE
'
'                Call PH_PY034_MTX01
'
''                oForm.Items("GCode").Click ct_Regular
'            End If
'
'        ElseIf pval.ItemUID = "Btn01" Then '/조회
'
'            oForm.Mode = fm_ADD_MODE '/fm_VIEW_MODE
'
'            Call PH_PY034_MTX01
'
'        End If
'
'    ElseIf pval.BeforeAction = False Then
'        If pval.ItemUID = "PH_PY034" Then
'            If oForm.Mode = fm_ADD_MODE Then
'            ElseIf oForm.Mode = fm_UPDATE_MODE Then
'            ElseIf oForm.Mode = fm_OK_MODE Then
'            End If
'        End If
'    End If
    
    Exit Sub
    
Raise_EVENT_ITEM_PRESSED_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ITEM_PRESSED_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_KEY_DOWN(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_KEY_DOWN_Error

    If pval.BeforeAction = True Then
    
        Call MDC_PS_Common.ActiveUserDefineValue(oForm, pval, BubbleEvent, "TeamCode", "") '팀정보
        Call MDC_PS_Common.ActiveUserDefineValue(oForm, pval, BubbleEvent, "Mat01", "ProfCode") '배부규칙
        
    ElseIf pval.BeforeAction = False Then
    
    End If
    
    Exit Sub
    
Raise_EVENT_KEY_DOWN_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_KEY_DOWN_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_CLICK(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_CLICK_Error

    If pval.BeforeAction = True Then
    
        If pval.ItemUID = "Mat01" Then
            If pval.Row > 0 Then
                oLast_Item_UID = pval.ItemUID
                oLast_Col_UID = pval.ColUID
                oLast_Col_Row = pval.Row
                
                Call oMat01.SelectRow(pval.Row, True, False)
            End If
        Else
            oLast_Item_UID = pval.ItemUID
            oLast_Col_UID = ""
            oLast_Col_Row = 0
        End If
    
    ElseIf pval.BeforeAction = False Then
    
    End If
    
    Exit Sub
    
Raise_EVENT_CLICK_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_CLICK_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_COMBO_SELECT(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_COMBO_SELECT_Error

    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
    
'        Call PH_PY034_FlushToItemValue(pval.ItemUID)
        
    End If

    Exit Sub
Raise_EVENT_COMBO_SELECT_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_COMBO_SELECT_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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

    Call oForm.Freeze(True)
    
    If pval.BeforeAction = True Then
    
        If pval.ItemChanged = True Then
        
            Call PH_PY034_FlushToItemValue(pval.ItemUID)
        
'            If (pval.ItemUID = "Mat01") Then
''                If (pval.ColUID = "ItemCode") Then
''                    '//기타작업
''                    Call oDS_PH_PY032B.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
''                    If oMat01.RowCount = pval.Row And Trim(oDS_PH_PY032B.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
''                        PH_PY032_AddMatrixRow (pval.Row)
''                    End If
''                Else
''                    Call oDS_PH_PY032B.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
''                End If
'            Else
'
'                Call PH_PY034_FlushToItemValue(pval.ItemUID)
'
'            End If
'            oMat01.LoadFromDataSource
'            oMat01.AutoResizeColumns
'            oForm.Update
        End If
        
    ElseIf pval.BeforeAction = False Then

    End If
    
    Call oForm.Freeze(False)
    
    Exit Sub
    
Raise_EVENT_VALIDATE_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_VALIDATE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_MATRIX_LOAD(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_MATRIX_LOAD_Error

    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
        Call PH_PY034_FormItemEnabled
        Call PH_PY034_AddMatrixRow   '//UDO방식
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
        Call PH_PY034_FormResize
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
'            oForm.DataSources.UserDataSources("ItemCode").Value = oDataTable01.Columns(0).Cells(0).Value
'            Set oDataTable01 = Nothing
'        End If
'        If (pval.ItemUID = "CardCode" Or pval.ItemUID = "CardName") Then
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY032A", "U_CardCode,U_CardName")
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
            oLast_Item_UID = pval.ItemUID
            oLast_Col_UID = pval.ColUID
            oLast_Col_Row = pval.Row
        End If
    Else
        oLast_Item_UID = pval.ItemUID
        oLast_Col_UID = ""
        oLast_Col_Row = 0
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
        Set oForm = Nothing
        Set oMat01 = Nothing
    End If
    
    Exit Sub
    
Raise_EVENT_FORM_UNLOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_FORM_UNLOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_ROW_DELETE(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_ROW_DELETE_Error

    Dim i As Long
    
    If (oLast_Col_Row > 0) Then
        If pval.BeforeAction = True Then
'            If (PH_PY032_Validate("행삭제") = False) Then
'                BubbleEvent = False
'                Exit Sub
'            End If
            '//행삭제전 행삭제가능여부검사
        ElseIf pval.BeforeAction = False Then
            For i = 1 To oMat01.VisualRowCount
                oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
            Next i
            oMat01.FlushToDataSource
            Call oDS_PH_PY034B.RemoveRecord(oDS_PH_PY034B.Size - 1)
            oMat01.LoadFromDataSource
            If oMat01.RowCount = 0 Then
                Call PH_PY034_AddMatrixRow
            Else
                If Trim(oDS_PH_PY034B.GetValue("U_DocNo", oMat01.RowCount - 1)) <> "" Then
                    Call PH_PY034_AddMatrixRow
                End If
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PH_PY034_CreateItems() As Boolean
On Error GoTo PH_PY034_CreateItems_Error

    Call oForm.Freeze(True)
    
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Set oDS_PH_PY034A = oForm.DataSources.DBDataSources("@PH_PY034A")
    Set oDS_PH_PY034B = oForm.DataSources.DBDataSources("@PH_PY034B")
    
    '// 메트릭스 개체 할당
    Set oMat01 = oForm.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
'    '관리번호
'    Call oForm.DataSources.UserDataSources.Add("SDocEntry", dt_SHORT_TEXT, 100)
'    Call oForm.Items("SDocEntry").Specific.DataBind.SetBound(True, "", "SDocEntry")
'
'    '사업장
'    Call oForm.DataSources.UserDataSources.Add("SCLTCOD", dt_SHORT_TEXT, 10)
'    Call oForm.Items("SCLTCOD").Specific.DataBind.SetBound(True, "", "SCLTCOD")
'
'    '부서
'    Call oForm.DataSources.UserDataSources.Add("STeamCd", dt_SHORT_TEXT, 10)
'    Call oForm.Items("STeamCd").Specific.DataBind.SetBound(True, "", "STeamCd")
'
'    '사원번호
'    Call oForm.DataSources.UserDataSources.Add("SMSTCOD", dt_SHORT_TEXT, 20)
'    Call oForm.Items("SMSTCOD").Specific.DataBind.SetBound(True, "", "SMSTCOD")
'
'    '사원성명
'    Call oForm.DataSources.UserDataSources.Add("SMSTNAM", dt_SHORT_TEXT, 50)
'    Call oForm.Items("SMSTNAM").Specific.DataBind.SetBound(True, "", "SMSTNAM")
'
'    '시작월
'    Call oForm.DataSources.UserDataSources.Add("SFrDate", dt_SHORT_TEXT, 10)
'    Call oForm.Items("SFrDate").Specific.DataBind.SetBound(True, "", "SFrDate")
'
'    '종료월
'    Call oForm.DataSources.UserDataSources.Add("SToDate", dt_SHORT_TEXT, 10)
'    Call oForm.Items("SToDate").Specific.DataBind.SetBound(True, "", "SToDate")
'
'    '목적
'    Call oForm.DataSources.UserDataSources.Add("SObject", dt_LONG_TEXT, 100)
'    Call oForm.Items("SObject").Specific.DataBind.SetBound(True, "", "SObject")
    
    Set oRecordSet01 = Nothing
    Call oForm.Freeze(False)
    
    Exit Function
PH_PY034_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY034_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY034_ComboBox_Setting()
On Error GoTo PH_PY034_ComboBox_Setting_Error

    Dim oCombo As SAPbouiCOM.ComboBox
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Call oForm.Freeze(True)
 
    '//////////조회정보//////////
'    '부서
'    Call oForm.Items("STeamCd").Specific.ValidValues.Add("%", "전체")
'    sQry = "            SELECT      U_Code AS [Code],"
'    sQry = sQry & "                 U_CodeNm As [Name]"
'    sQry = sQry & "  FROM       [@PS_HR200L]"
'    sQry = sQry & "  WHERE      Code = '1'"
'    sQry = sQry & "                 AND U_UseYN = 'Y'"
'    sQry = sQry & "  ORDER BY  U_Seq"
'    Call MDC_SetMod.Set_ComboList(oForm.Items("STeamCd").Specific, sQry, "", False, False)
'    Call oForm.Items("STeamCd").Specific.Select(0, psk_Index)

    '//////////매트릭스//////////
    '사업장
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("CLTCOD"), "SELECT BPLId, BPLName FROM OBPL order by BPLId")
    
    Call oForm.Freeze(False)
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    
    Exit Sub
PH_PY034_ComboBox_Setting_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY034_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY034_CF_ChooseFromList()
On Error GoTo PH_PY034_CF_ChooseFromList_Error

    Exit Sub
PH_PY034_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY034_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY034_EnableMenus()
On Error GoTo PH_PY034_EnableMenus_Error
    
    Call oForm.EnableMenu("1283", False) '삭제
    Call oForm.EnableMenu("1286", False) '닫기(미지원)
    Call oForm.EnableMenu("1287", False) '복제
    Call oForm.EnableMenu("1285", False) '복원
    Call oForm.EnableMenu("1284", True) '취소
    Call oForm.EnableMenu("1293", False) '행삭제(미지원)
    Call oForm.EnableMenu("1281", False)
    Call oForm.EnableMenu("1282", True)
    
    Exit Sub
    
PH_PY034_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY034_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY034_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY034_SetDocument_Error

    If (oFromDocEntry01 = "") Then
        Call PH_PY034_FormItemEnabled
        Call PH_PY034_AddMatrixRow
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY034_FormItemEnabled
        oForm.Items("DocEntry").Specific.VALUE = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub

PH_PY034_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY034_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY034_FormResize()
On Error GoTo PH_PY034_FormResize_Error
    
    Call oMat01.AutoResizeColumns
    
    Exit Sub
PH_PY034_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY034_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

'/폼의 아이템 사용지정
Private Sub PH_PY034_FormItemEnabled()
On Error GoTo PH_PY034_FormItemEnabled_Error
    If (oForm.Mode = fm_ADD_MODE) Then
        
        oForm.Items("CLTCOD").Enabled = True
        oForm.Items("DocEntry").Enabled = False
        oForm.Items("TeamCode").Enabled = True
        oForm.Items("FrDt").Enabled = True
        oForm.Items("ToDt").Enabled = True
        oForm.Items("JdtDate").Enabled = True
        
        oForm.Items("Mat01").Enabled = True
    
        oForm.Items("Btn02").Enabled = True
    
        '폼 DocEntry 세팅
        Call PH_PY034_FormClear
    
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        
        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", False)    '//문서추가
        
    ElseIf (oForm.Mode = fm_FIND_MODE) Then
    
        oForm.Items("CLTCOD").Enabled = True
        oForm.Items("DocEntry").Enabled = True
        oForm.Items("TeamCode").Enabled = True
        oForm.Items("FrDt").Enabled = True
        oForm.Items("ToDt").Enabled = True
        oForm.Items("JdtDate").Enabled = True
        
        oForm.Items("Mat01").Enabled = True
    
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        
        Call oForm.EnableMenu("1281", False)    '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가
        
    ElseIf (oForm.Mode = fm_OK_MODE) Then
    
        oForm.Items("CLTCOD").Enabled = False
        oForm.Items("DocEntry").Enabled = False
        oForm.Items("TeamCode").Enabled = False
        oForm.Items("FrDt").Enabled = False
        oForm.Items("ToDt").Enabled = False
        
'        If oForm.Items("JdtDate").Specific.VALUE <> "" Then '분개일자가 등록이 되어 있는 경우
            oForm.Items("JdtDate").Enabled = True
'        Else
'            oForm.Items("JdtDate").Enabled = False
'        End If
        
        oForm.Items("Mat01").Enabled = False
    
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        
        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가
        
    End If
    
    Exit Sub
    
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY034_FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "PH_PY034_FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub PH_PY034_FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo PH_PY034_FlushToItemValue_Error

    Dim i As Integer
    Dim ErrNum As Integer
    Dim sQry As String
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Dim CLTCOD As String
    
    Select Case oUID
        
        Case "Mat01"
            
            Call oMat01.AutoResizeColumns
        
        Case "CLTCOD"
            
        Case "TeamCode"
                
            oForm.Items("TeamName").Specific.VALUE = MDC_GetData.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" & oForm.Items("TeamCode").Specific.VALUE & "'", " AND Code = '1'") '팀명
        
    End Select

   Set oRecordSet01 = Nothing
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Exit Sub
PH_PY034_FlushToItemValue_Error:

    Set oRecordSet01 = Nothing

    Call MDC_Com.MDC_GF_Message("PH_PY034_FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E")

End Sub

Private Sub PH_PY034_AddMatrixRow()
    Dim oRow    As Long
    
    On Error GoTo PH_PY034_AddMatrixRow_Error
    
    Call oForm.Freeze(True)
    
    '//[Mat1]
    oMat01.FlushToDataSource
    oRow = oMat01.VisualRowCount
    
    If oMat01.VisualRowCount > 0 Then
        If Trim(oDS_PH_PY034B.GetValue("U_DocNo", oRow - 1)) <> "" Then
            If oDS_PH_PY034B.Size <= oMat01.VisualRowCount Then
                Call oDS_PH_PY034B.InsertRecord(oRow)
            End If
            oDS_PH_PY034B.Offset = oRow
            Call oDS_PH_PY034B.setValue("U_LineNum", oRow, oRow + 1)
            Call oDS_PH_PY034B.setValue("U_DocNo", oRow, "")
            Call oDS_PH_PY034B.setValue("U_CLTCOD", oRow, "")
            Call oDS_PH_PY034B.setValue("U_DestNo1", oRow, "")
            Call oDS_PH_PY034B.setValue("U_DestNo2", oRow, "")
            Call oDS_PH_PY034B.setValue("U_AcctCode", oRow, "")
            Call oDS_PH_PY034B.setValue("U_AcctName", oRow, "")
            Call oDS_PH_PY034B.setValue("U_Debit", oRow, 0)
            Call oDS_PH_PY034B.setValue("U_Credit", oRow, 0)
            Call oDS_PH_PY034B.setValue("U_ProfCode", oRow, "")
            Call oDS_PH_PY034B.setValue("U_Spender", oRow, "")
            Call oDS_PH_PY034B.setValue("U_LineMemo", oRow, "")
            Call oMat01.LoadFromDataSource
        Else
            oDS_PH_PY034B.Offset = oRow - 1
            Call oDS_PH_PY034B.setValue("U_LineNum", oRow - 1, oRow)
            Call oDS_PH_PY034B.setValue("U_DocNo", oRow - 1, "")
            Call oDS_PH_PY034B.setValue("U_CLTCOD", oRow - 1, "")
            Call oDS_PH_PY034B.setValue("U_DestNo1", oRow - 1, "")
            Call oDS_PH_PY034B.setValue("U_DestNo2", oRow - 1, "")
            Call oDS_PH_PY034B.setValue("U_AcctCode", oRow - 1, "")
            Call oDS_PH_PY034B.setValue("U_AcctName", oRow - 1, "")
            Call oDS_PH_PY034B.setValue("U_Debit", oRow - 1, 0)
            Call oDS_PH_PY034B.setValue("U_Credit", oRow - 1, 0)
            Call oDS_PH_PY034B.setValue("U_ProfCode", oRow - 1, "")
            Call oDS_PH_PY034B.setValue("U_Spender", oRow, "")
            Call oDS_PH_PY034B.setValue("U_LineMemo", oRow - 1, "")
            Call oMat01.LoadFromDataSource
        End If
    ElseIf oMat01.VisualRowCount = 0 Then
        oDS_PH_PY034B.Offset = oRow
        Call oDS_PH_PY034B.setValue("U_LineNum", oRow, oRow + 1)
        Call oDS_PH_PY034B.setValue("U_DocNo", oRow, "")
        Call oDS_PH_PY034B.setValue("U_CLTCOD", oRow, "")
        Call oDS_PH_PY034B.setValue("U_DestNo1", oRow, "")
        Call oDS_PH_PY034B.setValue("U_DestNo2", oRow, "")
        Call oDS_PH_PY034B.setValue("U_AcctCode", oRow, "")
        Call oDS_PH_PY034B.setValue("U_AcctName", oRow, "")
        Call oDS_PH_PY034B.setValue("U_Debit", oRow, 0)
        Call oDS_PH_PY034B.setValue("U_Credit", oRow, 0)
        Call oDS_PH_PY034B.setValue("U_ProfCode", oRow, "")
        Call oDS_PH_PY034B.setValue("U_Spender", oRow, "")
        Call oDS_PH_PY034B.setValue("U_LineMemo", oRow, "")
        Call oMat01.LoadFromDataSource
    End If
    
    Call oForm.Freeze(False)
    Exit Sub
PH_PY034_AddMatrixRow_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY034_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY034_FormClear()
On Error GoTo PH_PY034_FormClear_Error

    Dim DocEntry As String
    
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY034'", "")
    
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    
    Exit Sub
    
PH_PY034_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY034_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PH_PY034_DataValidCheck() As Boolean
On Error GoTo PH_PY034_DataValidCheck_Error

    PH_PY034_DataValidCheck = False
    
    Dim i    As Long
    Dim sQry As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    '사업장
    If Trim(oDS_PH_PY034A.GetValue("U_CLTCOD", 0)) = "" Then
        Call Sbo_Application.SetStatusBarMessage("사업장은 필수입니다.", bmt_Short, True)
        Call oForm.Items("CLTCOD").CLICK(ct_Regular)
        PH_PY034_DataValidCheck = False
        Exit Function
    End If
    
    '팀
    If Trim(oDS_PH_PY034A.GetValue("U_TeamCode", 0)) = "" Then
        Call Sbo_Application.SetStatusBarMessage("팀정보는 필수입니다.", bmt_Short, True)
        Call oForm.Items("TeamCode").CLICK(ct_Regular)
        PH_PY034_DataValidCheck = False
        Exit Function
    End If
    
    '기간(시작)
    If Trim(oDS_PH_PY034A.GetValue("U_FrDt", 0)) = "" Then
        Call Sbo_Application.SetStatusBarMessage("기간(시작) 필수입니다.", bmt_Short, True)
        Call oForm.Items("FrDt").CLICK(ct_Regular)
        PH_PY034_DataValidCheck = False
        Exit Function
    End If
    
    '기간(종료)
    If Trim(oDS_PH_PY034A.GetValue("U_ToDt", 0)) = "" Then
        Call Sbo_Application.SetStatusBarMessage("기간(종료)는 필수입니다.", bmt_Short, True)
        Call oForm.Items("ToDt").CLICK(ct_Regular)
        PH_PY034_DataValidCheck = False
        Exit Function
    End If
    
    '라인
    If oMat01.VisualRowCount > 1 Then
        For i = 1 To oMat01.VisualRowCount - 1
        
            '배부규칙
            If oMat01.Columns("ProfCode").Cells(i).Specific.VALUE = "" And oMat01.Columns("DocNo").Cells(i).Specific.VALUE <> 0 Then
                Call Sbo_Application.SetStatusBarMessage("배부규칙은 필수입니다.", bmt_Short, True)
                Call oMat01.Columns("ProfCode").Cells(i).CLICK(ct_Regular)
                PH_PY034_DataValidCheck = False
                Exit Function
            End If
            
            '적요
            If oMat01.Columns("LineMemo").Cells(i).Specific.VALUE = "" And oMat01.Columns("DocNo").Cells(i).Specific.VALUE <> 0 Then
                Call Sbo_Application.SetStatusBarMessage("적요는 필수입니다.", bmt_Short, True)
                Call oMat01.Columns("LineMemo").Cells(i).CLICK(ct_Regular)
                PH_PY034_DataValidCheck = False
                Exit Function
            End If
            
        Next
    Else
        Sbo_Application.SetStatusBarMessage "라인 데이터가 없습니다.", bmt_Short, True
        PH_PY034_DataValidCheck = False
        Exit Function
    End If
    
        oMat01.FlushToDataSource
    '// Matrix 마지막 행 삭제(DB 저장시)
    If oDS_PH_PY034B.Size > 1 Then oDS_PH_PY034B.RemoveRecord (oDS_PH_PY034B.Size - 1)
    
    oMat01.LoadFromDataSource
    
    PH_PY034_DataValidCheck = True
    
    Exit Function
    
    Set oRecordSet = Nothing
    
PH_PY034_DataValidCheck_Error:
    Set oRecordSet = Nothing
    PH_PY034_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY034_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Public Sub PH_PY034_MTX01()
'******************************************************************************
'Function ID : PH_PY034_MTX01()
'해당모듈    : PH_PY034
'기능        : 데이터 조회
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PH_PY034_MTX01_Error

    Dim i      As Integer
    Dim sQry   As String
    Dim ErrNum As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Dim CLTCOD   As String '사업장
    Dim TeamCode As String '팀
    Dim FrDt     As String '기간(시작)
    Dim ToDt     As String '기간(종료)
    
    CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE) '사업장
    TeamCode = Trim(oForm.Items("TeamCode").Specific.VALUE) '부서
    FrDt = Replace(Trim(oForm.Items("FrDt").Specific.VALUE), ".", "") '시작일자
    ToDt = Replace(Trim(oForm.Items("ToDt").Specific.VALUE), ".", "") '종료일자
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    Call oForm.Freeze(True)
    
    sQry = "            EXEC [PH_PY034_01] "
    sQry = sQry & "'" & CLTCOD & "',"  '사업장
    sQry = sQry & "'" & TeamCode & "',"  '부서
    sQry = sQry & "'" & FrDt & "',"  '시작일자
    sQry = sQry & "'" & ToDt & "'"  '종료일자
    
    Call oRecordSet01.DoQuery(sQry)
    
    oMat01.Clear
    oDS_PH_PY034B.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource

    If (oRecordSet01.RecordCount = 0) Then
    
        ErrNum = 1
        
'        oForm.Mode = fm_ADD_MODE
        
        Call PH_PY034_FormItemEnabled
        Call PH_PY034_AddMatrixRow
        
        GoTo PH_PY034_MTX01_Error
        
        Exit Sub
    End If
    
    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PH_PY034B.Size Then
            Call oDS_PH_PY034B.InsertRecord(i)
        End If
        
        oMat01.AddRow
        oDS_PH_PY034B.Offset = i
        
        Call oDS_PH_PY034B.setValue("U_LineNum", i, i + 1)
        Call oDS_PH_PY034B.setValue("U_DocNo", i, Trim(oRecordSet01.Fields("DocNo").VALUE)) '관리번호
        Call oDS_PH_PY034B.setValue("U_CLTCOD", i, Trim(oRecordSet01.Fields("CLTCOD").VALUE)) '사업장
        Call oDS_PH_PY034B.setValue("U_DestNo1", i, Trim(oRecordSet01.Fields("DestNo1").VALUE)) '출장번호1
        Call oDS_PH_PY034B.setValue("U_DestNo2", i, Trim(oRecordSet01.Fields("DestNo2").VALUE)) '출장번호2
        Call oDS_PH_PY034B.setValue("U_AcctCode", i, Trim(oRecordSet01.Fields("AcctCode").VALUE)) 'G/L계정
        Call oDS_PH_PY034B.setValue("U_AcctName", i, Trim(oRecordSet01.Fields("AcctName").VALUE)) 'G/L계정명
        Call oDS_PH_PY034B.setValue("U_Debit", i, Trim(oRecordSet01.Fields("Debit").VALUE)) '차변
        Call oDS_PH_PY034B.setValue("U_Credit", i, Trim(oRecordSet01.Fields("Credit").VALUE)) '대변
        Call oDS_PH_PY034B.setValue("U_ProfCode", i, Trim(oRecordSet01.Fields("ProfCode").VALUE)) '배부규칙
        Call oDS_PH_PY034B.setValue("U_Spender", i, Trim(oRecordSet01.Fields("Spender").VALUE)) '사용인
        Call oDS_PH_PY034B.setValue("U_LineMemo", i, Trim(oRecordSet01.Fields("LineMemo").VALUE)) '적요
        
        oRecordSet01.MoveNext
        ProgBar01.VALUE = ProgBar01.VALUE + 1
        ProgBar01.Text = ProgBar01.VALUE & "/" & oRecordSet01.RecordCount & "건 조회중...!"
        
    Next
    
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    Call PH_PY034_AddMatrixRow
    ProgBar01.Stop
    Call oForm.Freeze(False)
  
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PH_PY034_MTX01_Error:
'    ProgBar01.Stop
    Call oForm.Freeze(False)
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("조회 결과가 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PH_PY034_MTX01_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub


'Private Function Exist_YN(DocNum$) As String
'
'    Dim oRecordSet01   As SAPbobsCOM.Recordset
'    Dim sQry           As String
'
'    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
'
'    sQry = "SELECT Top 1 T1.DocNum FROM [@PH_PY034H] T1 "
'    sQry = sQry & " WHERE T1.DocNum  = '" & DocNum & "'"
'    oRecordSet01.DoQuery sQry
'
'    Do Until oRecordSet01.EOF
'        Exist_YN = oRecordSet01(0).VALUE
'        oRecordSet01.MoveNext
'    Loop
'
'    If Trim(Exist_YN) = "" Then
'        Exist_YN = ""
'        Exit Function
'    End If
'
'    Set oRecordSet01 = Nothing
'End Function

Private Function PH_PY034_Item_Error_Message(ErrNum As Integer)
On Error GoTo PH_PY034_Item_Error_Message

    GoTo PH_PY034_Item_Error_Message

PH_PY034_Item_Error_Message:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "분개처리일을 먼저 입력하세요.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "문서가 Close 또는 Cancel 되었습니다.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "분개생성:Y일 때 취소 할 수 있습니다.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "거래처코드와, 사업장을 먼저 입력하세요.", "E"
 End If
End Function

Private Function PH_PY034_Create_oJournalEntries(ChkType As Integer) As Boolean
'******************************************************************************
'Function ID : PH_PY034_Create_oJournalEntries()
'해당모듈    : PH_PY034
'기능        : 분개 생성
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PH_PY034_Create_oJournalEntries

    Dim f_oJournalEntries As SAPbobsCOM.JournalEntries
    
    Dim i        As Long
    Dim ErrNum   As Integer
    Dim ErrCode  As Long
    Dim ErrMsg   As String
    Dim RetVal   As String
    Dim sTransId As String
    Dim sDebit   As Currency
    Dim sCredit  As Currency

    Dim sAcctCode As String
    Dim sDocDate  As String
    Dim sPrcCode  As String
    Dim sSpender  As String '사용인
    Dim sLineMemo As String

    Dim sCC As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Dim sQry As String

    If (oCompany.InTransaction = True) Then
        oCompany.EndTransaction wf_RollBack
    End If
    oCompany.StartTransaction

    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)

    oMat01.FlushToDataSource

    sDocDate = Format(oDS_PH_PY034A.GetValue("U_JdtDate", 0), "0000-00-00")

    Set f_oJournalEntries = oCompany.GetBusinessObject(oJournalEntries)

    With f_oJournalEntries
         .ReferenceDate = sDocDate            '전기일
         .DueDate = sDocDate
         .TaxDate = sDocDate

         For i = 1 To oMat01.VisualRowCount
            sAcctCode = oMat01.Columns("AcctCode").Cells(i).Specific.VALUE 'G/L 계정
            sDebit = oMat01.Columns("Debit").Cells(i).Specific.VALUE '차변
            sCredit = oMat01.Columns("Credit").Cells(i).Specific.VALUE '대변
            sPrcCode = oMat01.Columns("ProfCode").Cells(i).Specific.VALUE '배부규칙
            sSpender = oMat01.Columns("Spender").Cells(i).Specific.VALUE '사용인
            sLineMemo = oMat01.Columns("LineMemo").Cells(i).Specific.VALUE '적요
            
            .Lines.Add
            
            If sAcctCode <> "" Then

                .Lines.SetCurrentLine i - 1
                .Lines.AccountCode = sAcctCode '관리계정
                .Lines.ShortName = sAcctCode 'G/L계정/BP 코드
                .Lines.Debit = sDebit '차변
                .Lines.Credit = sCredit '대변
                .Lines.CostingCode = sPrcCode '배부규칙
                .Lines.LineMemo = sLineMemo '비고
                
                .Lines.UserFields("U_spender").VALUE = sSpender '사용인
                .Lines.UserFields("U_BillCode").VALUE = "P90010" '법정지출증빙코드
                .Lines.UserFields("U_BillName").VALUE = "규정" '법정지출증빙명
                
                .UserFields("U_BPLId").VALUE = Trim(oForm.Items("CLTCOD").Specific.VALUE) '사업장

            End If
            
        Next

    End With
    '// 완료
    RetVal = f_oJournalEntries.Add
    If (0 <> RetVal) Then
        Call oCompany.GetLastError(ErrCode, ErrMsg)
        GoTo PH_PY034_Create_oJournalEntries
    End If

    sCC = "Y"

    If ChkType = 1 Then
        Call oCompany.GetNewObjectCode(sTransId)
        sQry = "        UPDATE  [@PH_PY034A] "
        sQry = sQry & " SET     U_JdtNo = '" & sTransId & "',"
        sQry = sQry & "         U_JdtDate = '" & sDocDate & "',"
        sQry = sQry & "         U_JdtCC = '" & sCC & "'"
        sQry = sQry & " WHERE   DocNum = '" & Trim(oDS_PH_PY034A.GetValue("DocNum", 0)) & "'"
        
        Call oRecordSet01.DoQuery(sQry)
        
        If (oCompany.InTransaction = True) Then
            oCompany.EndTransaction wf_Commit
        End If
        
    End If

    Call oDS_PH_PY034A.setValue("U_JdtNo", 0, sTransId)
    Call oDS_PH_PY034A.setValue("U_JdtCC", 0, sCC)

    Set oRecordSet01 = Nothing
    Set f_oJournalEntries = Nothing
    PH_PY034_Create_oJournalEntries = True

    oForm.Items("Btn02").Enabled = False
    oForm.Items("Btn03").Enabled = True
    
    Call MDC_Com.MDC_GF_Message("성공적으로 분개등록되었습니다.", "S")

Exit Function

'///////////////////////////////////////////////////////////////////////////////////////
PH_PY034_Create_oJournalEntries:
    Set oRecordSet01 = Nothing
    Set f_oJournalEntries = Nothing
    If oCompany.InTransaction Then
        oCompany.EndTransaction wf_RollBack
    End If
    PH_PY034_Create_oJournalEntries = False
    MDC_Com.MDC_GF_Message "PH_PY034_Create_oJournalEntries : " & Err.Description & ErrMsg, "E"
End Function

Private Function PH_PY034_Cancel_oJournalEntries(ChkType As Integer) As Boolean
'******************************************************************************
'Function ID : PH_PY034_Cancel_oJournalEntries()
'해당모듈    : PH_PY034
'기능        : 분개 취소
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PH_PY034_Cancel_oJournalEntries

    Dim f_oJournalEntries As SAPbobsCOM.JournalEntries

    Dim i        As Long
    Dim ErrNum   As Integer
    Dim ErrCode  As Long
    Dim ErrMsg   As String
    Dim RetVal   As Long
    Dim sTransId As String

    Dim sCardCode As String
    Dim sDocDate  As String
    Dim sCC       As String

    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Dim sQry As String

    If (oCompany.InTransaction = True) Then
        oCompany.EndTransaction wf_RollBack
    End If

    oCompany.StartTransaction

    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)

    oMat01.FlushToDataSource

    Set f_oJournalEntries = Nothing
    Set f_oJournalEntries = oCompany.GetBusinessObject(oJournalEntries)

    If (f_oJournalEntries.GetByKey(Trim(oDS_PH_PY034A.GetValue("U_JdtNo", 0))) = False) Then
        Call oCompany.GetLastError(ErrCode, ErrMsg)
        GoTo PH_PY034_Cancel_oJournalEntries
    End If

    '// 완료
    RetVal = f_oJournalEntries.Cancel
    If (0 <> RetVal) Then
        Call oCompany.GetLastError(ErrCode, ErrMsg)
        GoTo PH_PY034_Cancel_oJournalEntries
    End If

    sCC = "N"

    If ChkType = 1 Then
        oCompany.GetNewObjectCode sTransId
        sQry = "        UPDATE  [@PH_PY034A]"
        sQry = sQry & " SET     U_JdtCanNo = '" & sTransId & "',"
        sQry = sQry & "         U_JdtCC = '" & sCC & "'"
        sQry = sQry & " WHERE   DocNum = '" & Trim(oDS_PH_PY034A.GetValue("DocNum", 0)) & "'"
        oRecordSet01.DoQuery sQry

        If (oCompany.InTransaction = True) Then
            oCompany.EndTransaction wf_Commit
        End If
    End If

'    oForm.Update
    Call oDS_PH_PY034A.setValue("U_JdtCanNo", 0, sTransId)
    Call oDS_PH_PY034A.setValue("U_JdtCC", 0, sCC)

    Set oRecordSet01 = Nothing
    Set f_oJournalEntries = Nothing
    PH_PY034_Cancel_oJournalEntries = True

    oForm.Items("Btn02").Enabled = False
    oForm.Items("Btn03").Enabled = False
    
    Call MDC_Com.MDC_GF_Message("성공적으로 분개취소되었습니다.", "S")

Exit Function

'///////////////////////////////////////////////////////////////////////////////////////
PH_PY034_Cancel_oJournalEntries:
    Set oRecordSet01 = Nothing
    Set f_oJournalEntries = Nothing
    If oCompany.InTransaction Then
        oCompany.EndTransaction wf_RollBack
    End If
    PH_PY034_Cancel_oJournalEntries = False
    MDC_Com.MDC_GF_Message "Create_oJournalEntries : " & Err.Description & ErrMsg, "E"
End Function




