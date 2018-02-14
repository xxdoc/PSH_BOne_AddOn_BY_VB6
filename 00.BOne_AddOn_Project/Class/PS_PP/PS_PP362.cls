VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_PP362"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File : PS_PP362.cls
'//  Module : 생산관리 > 애드온 > 생산관리
'//  Desc : 작번별 생산진행현황2
'//  FormType : PS_PP362
'//  Create Date(Start) : 2012.01.28
'//  Create Date(End) :
'//  Modified Date :
'//  Creator : Song Myoung gyu
'//  Company : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form

Public oMat01 As SAPbouiCOM.Matrix
Public oMat02 As SAPbouiCOM.Matrix
Public oMat03 As SAPbouiCOM.Matrix
Public oMat04 As SAPbouiCOM.Matrix
Public oMat05 As SAPbouiCOM.Matrix
Public oMat06 As SAPbouiCOM.Matrix
Public oMat07 As SAPbouiCOM.Matrix
Public oMat08 As SAPbouiCOM.Matrix

Private oDS_PS_PP362H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_PP362L As SAPbouiCOM.DBDataSource '라인(Main작번)
Private oDS_PS_PP362M As SAPbouiCOM.DBDataSource '라인(Sub작번)
Private oDS_PS_PP362N As SAPbouiCOM.DBDataSource '라인(자재비내역)
Private oDS_PS_PP362O As SAPbouiCOM.DBDataSource '라인(설계비내역)
Private oDS_PS_PP362P As SAPbouiCOM.DBDataSource '라인(자체가공비내역_공정별내역)
Private oDS_PS_PP362Q As SAPbouiCOM.DBDataSource '라인(자체가공비내역_작업자별내역)
Private oDS_PS_PP362R As SAPbouiCOM.DBDataSource '라인(외주가공비내역)
Private oDS_PS_PP362S As SAPbouiCOM.DBDataSource '라인(외주제작비내역)

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
'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01 As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_PP362.srf")
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
    
    oFormUniqueID01 = "PS_PP362_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    '//oForm01.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_PP362_CreateItems
    Call PS_PP362_ComboBox_Setting
    Call PS_PP362_Initial_Setting
    Call PS_PP362_CF_ChooseFromList
    Call PS_PP362_EnableMenus
    Call PS_PP362_SetDocument(oFromDocEntry01)
    Call PS_PP362_FormResize
    
    oForm01.Update
    oForm01.Freeze False
    
    oForm01.Items("Folder01").Specific.Select '폼이 로드 될 때 Folder01이 선택됨
    
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
            Case "1281": '찾기
            Case "1282": '추가
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
            
            Case "7169": '엑셀 내보내기
            
                '엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
                Call PS_PP362_AddMatrixRow(oMat01.VisualRowCount, oMat01, oDS_PS_PP362L) 'Main작번
                Call PS_PP362_AddMatrixRow(oMat02.VisualRowCount, oMat02, oDS_PS_PP362M) 'Sub작번
                Call PS_PP362_AddMatrixRow(oMat03.VisualRowCount, oMat03, oDS_PS_PP362N) '자재비내역
                Call PS_PP362_AddMatrixRow(oMat04.VisualRowCount, oMat04, oDS_PS_PP362O) '설계비내역
                Call PS_PP362_AddMatrixRow(oMat05.VisualRowCount, oMat05, oDS_PS_PP362P) '자체가공비내역_공정별
                Call PS_PP362_AddMatrixRow(oMat06.VisualRowCount, oMat06, oDS_PS_PP362Q) '자체가공비내역_작업자별
                Call PS_PP362_AddMatrixRow(oMat07.VisualRowCount, oMat07, oDS_PS_PP362R) '외주가공비내역
                Call PS_PP362_AddMatrixRow(oMat08.VisualRowCount, oMat08, oDS_PS_PP362S) '외주제작비내역
            
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
                'MAIN작번
                Call oDS_PS_PP362L.RemoveRecord(oDS_PS_PP362L.Size - 1)
                Call oMat01.LoadFromDataSource
                
                'Sub작번
                Call oDS_PS_PP362M.RemoveRecord(oDS_PS_PP362M.Size - 1)
                Call oMat02.LoadFromDataSource
                
                '자재비내역
                Call oDS_PS_PP362N.RemoveRecord(oDS_PS_PP362N.Size - 1)
                Call oMat03.LoadFromDataSource
                
                '설계비내역
                Call oDS_PS_PP362O.RemoveRecord(oDS_PS_PP362O.Size - 1)
                Call oMat04.LoadFromDataSource
                
                '자체가공비내역_공정별
                Call oDS_PS_PP362P.RemoveRecord(oDS_PS_PP362P.Size - 1)
                Call oMat05.LoadFromDataSource
                
                '자체가공비내역_작업자별
                Call oDS_PS_PP362Q.RemoveRecord(oDS_PS_PP362Q.Size - 1)
                Call oMat06.LoadFromDataSource
                
                '외주가공비내역
                Call oDS_PS_PP362R.RemoveRecord(oDS_PS_PP362R.Size - 1)
                Call oMat07.LoadFromDataSource
                
                '외주제작비내역
                Call oDS_PS_PP362S.RemoveRecord(oDS_PS_PP362S.Size - 1)
                Call oMat08.LoadFromDataSource
                
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
    ElseIf pval.ItemUID = "Mat05" Then
        If pval.Row > 0 Then
            oLastItemUID01 = pval.ItemUID
            oLastColUID01 = pval.ColUID
            oLastColRow01 = pval.Row
        End If
    ElseIf pval.ItemUID = "Mat06" Then
        If pval.Row > 0 Then
            oLastItemUID01 = pval.ItemUID
            oLastColUID01 = pval.ColUID
            oLastColRow01 = pval.Row
        End If
    ElseIf pval.ItemUID = "Mat07" Then
        If pval.Row > 0 Then
            oLastItemUID01 = pval.ItemUID
            oLastColUID01 = pval.ColUID
            oLastColRow01 = pval.Row
        End If
    ElseIf pval.ItemUID = "Mat08" Then
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
    
        If pval.ItemUID = "btnSearch" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If PS_PP362_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                Else
                    
                    oForm01.Items("MainJak").Specific.VALUE = ""
                    oForm01.Items("SubJak").Specific.VALUE = ""
                    
                    oMat01.Clear
                    oDS_PS_PP362L.Clear
                    
                    '다른 TAB의 정보 초기화
                    oMat02.Clear
                    oDS_PS_PP362M.Clear
                    
                    oMat03.Clear
                    oDS_PS_PP362N.Clear
                    
                    oMat04.Clear
                    oDS_PS_PP362O.Clear
                    
                    oMat05.Clear
                    oDS_PS_PP362P.Clear
                    
                    oMat06.Clear
                    oDS_PS_PP362Q.Clear
                    
                    oMat07.Clear
                    oDS_PS_PP362R.Clear
                    
                    oMat08.Clear
                    oDS_PS_PP362S.Clear
                    
                    oForm01.Items("Folder01").Specific.Select 'Main작번 TAB 선택
                    Call PS_PP362_MTX01 '매트릭스에 데이터 로드
                    
                End If
            End If

        ElseIf pval.ItemUID = "Btn01" Or pval.ItemUID = "Btn02" Or pval.ItemUID = "Btn03" Or pval.ItemUID = "Btn04" Or pval.ItemUID = "Btn05" Or pval.ItemUID = "Btn06" Or pval.ItemUID = "Btn07" Or pval.ItemUID = "Btn08" Then

            If oForm01.Mode = fm_ADD_MODE Then

                Call PS_PP362_Print_Report(pval.ItemUID)

            End If
            
        End If
        
    ElseIf pval.BeforeAction = False Then
    
        Call oForm01.Freeze(True)
    
        '폴더를 사용할 때는 필수 소스_S
        'Folder01이 선택되었을 때
        If pval.ItemUID = "Folder01" Then
            
            oForm01.PaneLevel = 1
            
        End If
        
        'Folder02가 선택되었을 때
        If pval.ItemUID = "Folder02" Then
            
            oForm01.PaneLevel = 2
            
        End If
        
        'Folder03이 선택되었을 때
        If pval.ItemUID = "Folder03" Then
            
            oForm01.PaneLevel = 3
            
        End If
        
        'Folder04이 선택되었을 때
        If pval.ItemUID = "Folder04" Then
            
            oForm01.PaneLevel = 4
            
        End If
        
        'Folder05이 선택되었을 때
        If pval.ItemUID = "Folder05" Then
            
            oForm01.PaneLevel = 5
            
        End If
        
        'Folder06이 선택되었을 때
        If pval.ItemUID = "Folder06" Then
            
            oForm01.PaneLevel = 6
            
        End If
        
        'Folder07이 선택되었을 때
        If pval.ItemUID = "Folder07" Then
            
            oForm01.PaneLevel = 7
            
        End If
        '폴더를 사용할 때는 필수 소스_E
    
        If pval.ItemUID = "PS_PP362" Then
        
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
            
        End If
        
        Call oForm01.Freeze(False)
    
    End If
    
    Exit Sub
Raise_EVENT_ITEM_PRESSED_Error:
    Call oForm01.Freeze(False)
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
        ElseIf pval.ItemUID = "Mat03" Then
            If pval.Row > 0 Then
                
                Call oMat03.SelectRow(pval.Row, True, False)
                
                oLastItemUID01 = pval.ItemUID
                oLastColUID01 = pval.ColUID
                oLastColRow01 = pval.Row
            End If
        ElseIf pval.ItemUID = "Mat04" Then
            If pval.Row > 0 Then

                Call oMat04.SelectRow(pval.Row, True, False)

                oLastItemUID01 = pval.ItemUID
                oLastColUID01 = pval.ColUID
                oLastColRow01 = pval.Row
            End If
        ElseIf pval.ItemUID = "Mat05" Then
            If pval.Row > 0 Then

                Call oMat05.SelectRow(pval.Row, True, False)

                oLastItemUID01 = pval.ItemUID
                oLastColUID01 = pval.ColUID
                oLastColRow01 = pval.Row
            End If
        ElseIf pval.ItemUID = "Mat06" Then
            If pval.Row > 0 Then

                Call oMat06.SelectRow(pval.Row, True, False)

                oLastItemUID01 = pval.ItemUID
                oLastColUID01 = pval.ColUID
                oLastColRow01 = pval.Row
            End If
        ElseIf pval.ItemUID = "Mat07" Then
            If pval.Row > 0 Then

                Call oMat07.SelectRow(pval.Row, True, False)

                oLastItemUID01 = pval.ItemUID
                oLastColUID01 = pval.ColUID
                oLastColRow01 = pval.Row
            End If
        ElseIf pval.ItemUID = "Mat08" Then
            If pval.Row > 0 Then

                Call oMat08.SelectRow(pval.Row, True, False)

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
    
        If pval.ItemUID = "Mat01" Then 'Main작번

            If pval.Row = 0 Then

                '정렬
                oMat01.Columns(pval.ColUID).TitleObject.Sortable = True
                oMat01.FlushToDataSource

            Else
            
                '다른 TAB의 정보 초기화
                oMat02.Clear
                oDS_PS_PP362M.Clear
                
                oMat03.Clear
                oDS_PS_PP362N.Clear
                
                oMat04.Clear
                oDS_PS_PP362O.Clear
                
                oMat05.Clear
                oDS_PS_PP362P.Clear
                
                oMat06.Clear
                oDS_PS_PP362Q.Clear
                
                oMat07.Clear
                oDS_PS_PP362R.Clear
                
                oMat08.Clear
                oDS_PS_PP362S.Clear

                oForm01.Items("Folder02").Specific.Select 'Sub작번 TAB 선택
                Call PS_PP362_MTX02(oMat01.Columns("OrdNum").Cells(pval.Row).Specific.VALUE)
                
            End If

        ElseIf pval.ItemUID = "Mat02" Then 'Sub작번

            If pval.Row = 0 Then
            
                '정렬
                oMat02.Columns(pval.ColUID).TitleObject.Sortable = True
                oMat02.FlushToDataSource
            
            Else
            
                oForm01.Items("MainJak").Specific.VALUE = oMat02.Columns("OrdNum").Cells(pval.Row).Specific.VALUE
                oForm01.Items("SubJak").Specific.VALUE = oMat02.Columns("OrdSub").Cells(pval.Row).Specific.VALUE
                
                '다른 TAB의 정보 초기화
                oMat03.Clear
                oDS_PS_PP362N.Clear
                
                oMat04.Clear
                oDS_PS_PP362O.Clear
                
                oMat05.Clear
                oDS_PS_PP362P.Clear
                
                oMat06.Clear
                oDS_PS_PP362Q.Clear
                
                oMat07.Clear
                oDS_PS_PP362R.Clear
                
                oMat08.Clear
                oDS_PS_PP362S.Clear
                
                '자재비
                Call PS_PP362_MTX03(oMat02.Columns("OrdNum").Cells(pval.Row).Specific.VALUE, oMat02.Columns("OrdSub").Cells(pval.Row).Specific.VALUE)
                
                '설계비
                Call PS_PP362_MTX04(oMat02.Columns("OrdNum").Cells(pval.Row).Specific.VALUE, oMat02.Columns("OrdSub").Cells(pval.Row).Specific.VALUE)
                
                '자체가공비_공정별
                Call PS_PP362_MTX05(oMat02.Columns("OrdNum").Cells(pval.Row).Specific.VALUE, oMat02.Columns("OrdSub").Cells(pval.Row).Specific.VALUE)
                
                '자체가공비_작업자별
                
                '외주가공비
                Call PS_PP362_MTX07(oMat02.Columns("OrdNum").Cells(pval.Row).Specific.VALUE, oMat02.Columns("OrdSub").Cells(pval.Row).Specific.VALUE)
                
                '외주제작비
                Call PS_PP362_MTX08(oMat02.Columns("OrdNum").Cells(pval.Row).Specific.VALUE, oMat02.Columns("OrdSub").Cells(pval.Row).Specific.VALUE)
                
            End If
        
        ElseIf pval.ItemUID = "Mat03" Then '자재비내역

            If pval.Row = 0 Then

                '정렬
                oMat03.Columns(pval.ColUID).TitleObject.Sortable = True
                oMat03.FlushToDataSource

            End If
            
        ElseIf pval.ItemUID = "Mat04" Then '설계비내역

            If pval.Row = 0 Then

                '정렬
                oMat04.Columns(pval.ColUID).TitleObject.Sortable = True
                oMat04.FlushToDataSource

            End If
        
        ElseIf pval.ItemUID = "Mat05" Then '자체가공비내역(공정별내역)

            If pval.Row = 0 Then

                '정렬
                oMat05.Columns(pval.ColUID).TitleObject.Sortable = True
                oMat05.FlushToDataSource

            Else

                oForm01.Items("CpCode").Specific.VALUE = oMat05.Columns("CpCode").Cells(pval.Row).Specific.VALUE
                Call PS_PP362_MTX06(oMat05.Columns("OrdNum").Cells(pval.Row).Specific.VALUE, oMat05.Columns("Sub1_2").Cells(pval.Row).Specific.VALUE, oMat05.Columns("CpCode").Cells(pval.Row).Specific.VALUE)

            End If
        
        ElseIf pval.ItemUID = "Mat06" Then '자체가공비내역(작업자별)

            If pval.Row = 0 Then

                '정렬
                oMat06.Columns(pval.ColUID).TitleObject.Sortable = True
                oMat06.FlushToDataSource

            End If
        
        ElseIf pval.ItemUID = "Mat07" Then '외주가공비내역

            If pval.Row = 0 Then

                '정렬
                oMat07.Columns(pval.ColUID).TitleObject.Sortable = True
                oMat07.FlushToDataSource

            End If
        
        ElseIf pval.ItemUID = "Mat08" Then '외주제작비내역

            If pval.Row = 0 Then

                '정렬
                oMat08.Columns(pval.ColUID).TitleObject.Sortable = True
                oMat08.FlushToDataSource

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

    If pval.BeforeAction = True Then
    
        If pval.ItemUID = "Mat02" Then

            Dim PP030 As Object
            Set PP030 = New PS_PP030

            Call PP030.LoadForm(oMat02.Columns("PP030HNo").Cells(pval.Row).Specific.VALUE)

            Set PP030 = Nothing

        ElseIf pval.ItemUID = "Mat06" Then

            Dim PP040 As Object
            Set PP040 = New PS_PP040

            Call PP040.LoadForm(oMat06.Columns("PP040HNo").Cells(pval.Row).Specific.VALUE)

            Set PP040 = Nothing

        End If
    
    ElseIf pval.BeforeAction = False Then
    
    End If
    
    Exit Sub
Raise_EVENT_MATRIX_LINK_PRESSED_Error:
'    Set PP030 = Nothing
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LINK_PRESSED_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
        Call PS_PP362_FormItemEnabled
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_PP362_FormResize
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
    ElseIf pval.ItemUID = "Mat05" Then
        If pval.Row > 0 Then
            oLastItemUID01 = pval.ItemUID
            oLastColUID01 = pval.ColUID
            oLastColRow01 = pval.Row
        End If
    ElseIf pval.ItemUID = "Mat06" Then
        If pval.Row > 0 Then
            oLastItemUID01 = pval.ItemUID
            oLastColUID01 = pval.ColUID
            oLastColRow01 = pval.Row
        End If
    ElseIf pval.ItemUID = "Mat07" Then
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
        Set oMat03 = Nothing
        Set oMat04 = Nothing
        Set oMat05 = Nothing
        Set oMat06 = Nothing
        Set oMat07 = Nothing
        Set oMat08 = Nothing
        
        Set oDS_PS_PP362H = Nothing
        Set oDS_PS_PP362L = Nothing
        Set oDS_PS_PP362M = Nothing
        Set oDS_PS_PP362N = Nothing
        Set oDS_PS_PP362O = Nothing
        Set oDS_PS_PP362P = Nothing
        Set oDS_PS_PP362Q = Nothing
        Set oDS_PS_PP362R = Nothing
        Set oDS_PS_PP362S = Nothing
        
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

Private Function PS_PP362_CreateItems() As Boolean
On Error GoTo PS_PP362_CreateItems_Error

    Call oForm01.Freeze(True)
    
'    Dim oQuery01 As String
'    Dim oRecordSet01 As SAPbobsCOM.Recordset
'    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Set oDS_PS_PP362L = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    Set oDS_PS_PP362M = oForm01.DataSources.DBDataSources("@PS_USERDS02")
    Set oDS_PS_PP362N = oForm01.DataSources.DBDataSources("@PS_USERDS03")
    Set oDS_PS_PP362O = oForm01.DataSources.DBDataSources("@PS_USERDS04")
    Set oDS_PS_PP362P = oForm01.DataSources.DBDataSources("@PS_USERDS05")
    Set oDS_PS_PP362Q = oForm01.DataSources.DBDataSources("@PS_USERDS06")
    Set oDS_PS_PP362R = oForm01.DataSources.DBDataSources("@PS_USERDS07")
    Set oDS_PS_PP362S = oForm01.DataSources.DBDataSources("@PS_USERDS08")
    
    '매트릭스 초기화
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    Set oMat02 = oForm01.Items("Mat02").Specific
    oMat02.SelectionMode = ms_NotSupported
    oMat02.AutoResizeColumns
    
    Set oMat03 = oForm01.Items("Mat03").Specific
    oMat03.SelectionMode = ms_NotSupported
    oMat03.AutoResizeColumns

    Set oMat04 = oForm01.Items("Mat04").Specific
    oMat04.SelectionMode = ms_NotSupported
    oMat04.AutoResizeColumns
    
    Set oMat05 = oForm01.Items("Mat05").Specific
    oMat05.SelectionMode = ms_NotSupported
    oMat05.AutoResizeColumns
    
    Set oMat06 = oForm01.Items("Mat06").Specific
    oMat06.SelectionMode = ms_NotSupported
    oMat06.AutoResizeColumns
    
    Set oMat07 = oForm01.Items("Mat07").Specific
    oMat07.SelectionMode = ms_NotSupported
    oMat07.AutoResizeColumns
    
    Set oMat08 = oForm01.Items("Mat08").Specific
    oMat08.SelectionMode = ms_NotSupported
    oMat08.AutoResizeColumns
    
    '작번등록년월(시작)_S
    Call oForm01.DataSources.UserDataSources.Add("FrDt", dt_SHORT_TEXT, 7)
    Call oForm01.Items("FrDt").Specific.DataBind.SetBound(True, "", "FrDt")
    oForm01.DataSources.UserDataSources.Item("FrDt").VALUE = Format(Now, "YYYY-MM")
    '작번등록년월(시작)_E
    
    '작번등록년월(종료)_S
    Call oForm01.DataSources.UserDataSources.Add("ToDt", dt_SHORT_TEXT, 7)
    Call oForm01.Items("ToDt").Specific.DataBind.SetBound(True, "", "ToDt")
    oForm01.DataSources.UserDataSources.Item("ToDt").VALUE = Format(Now, "YYYY-MM")
    '작번등록년월(종료)_E
    
    '품명_S
    Call oForm01.DataSources.UserDataSources.Add("FrgnName", dt_SHORT_TEXT, 100)
    Call oForm01.Items("FrgnName").Specific.DataBind.SetBound(True, "", "FrgnName")
    '품명_E
    
    '거래처구분_S
    Call oForm01.DataSources.UserDataSources.Add("CardType", dt_SHORT_TEXT, 10)
    Call oForm01.Items("CardType").Specific.DataBind.SetBound(True, "", "CardType")
    '거래처구분_E
    
    '규격_S
    Call oForm01.DataSources.UserDataSources.Add("Spec", dt_SHORT_TEXT, 100)
    Call oForm01.Items("Spec").Specific.DataBind.SetBound(True, "", "Spec")
    '규격_E
    
    '품목구분_S
    Call oForm01.DataSources.UserDataSources.Add("ItemClass", dt_SHORT_TEXT, 10)
    Call oForm01.Items("ItemClass").Specific.DataBind.SetBound(True, "", "ItemClass")
    '품목구분_E
    
    '생산완료여부_S
    Call oForm01.DataSources.UserDataSources.Add("WCYN", dt_SHORT_TEXT, 10)
    Call oForm01.Items("WCYN").Specific.DataBind.SetBound(True, "", "WCYN")
    '생산완료여부_E
    
    '거래처_S
    Call oForm01.DataSources.UserDataSources.Add("CardCode", dt_SHORT_TEXT, 50)
    Call oForm01.Items("CardCode").Specific.DataBind.SetBound(True, "", "CardCode")
    '거래처_E
    
    '거래처명_S
    Call oForm01.DataSources.UserDataSources.Add("CardName", dt_SHORT_TEXT, 100)
    Call oForm01.Items("CardName").Specific.DataBind.SetBound(True, "", "CardName")
    '거래처명_E
    
    '품목(작번)_S
    Call oForm01.DataSources.UserDataSources.Add("ItemCode", dt_SHORT_TEXT, 50)
    Call oForm01.Items("ItemCode").Specific.DataBind.SetBound(True, "", "ItemCode")
    '품목(작번)_E
    
    '품목명_S
    Call oForm01.DataSources.UserDataSources.Add("ItemName", dt_SHORT_TEXT, 100)
    Call oForm01.Items("ItemName").Specific.DataBind.SetBound(True, "", "ItemName")
    '품목명_E

    'Main작번_S
    Call oForm01.DataSources.UserDataSources.Add("MainJak", dt_SHORT_TEXT, 100)
    Call oForm01.Items("MainJak").Specific.DataBind.SetBound(True, "", "MainJak")
    'Main작번_E
    
    'Sub작번_S
    Call oForm01.DataSources.UserDataSources.Add("SubJak", dt_SHORT_TEXT, 100)
    Call oForm01.Items("SubJak").Specific.DataBind.SetBound(True, "", "SubJak")
    'Sub작번_E

    '공정코드_S
    Call oForm01.DataSources.UserDataSources.Add("CpCode", dt_SHORT_TEXT, 100)
    Call oForm01.Items("CpCode").Specific.DataBind.SetBound(True, "", "CpCode")
    '공정코드_E

    '자체가공비내역_공정별내역 Matrix의 필드 Hidden
    oMat05.Columns("ReTime").Visible = False '수정공수
    oMat05.Columns("CpCode").Visible = False '공정코드
    oMat05.Columns("OrdNum").Visible = False '작번
    oMat05.Columns("Sub1_2").Visible = False '서브작번
    oMat05.Columns("Class").Visible = False '공정완료여부(2012.08.06 송명규 수정)
    
    '자체가공비내역_작업자별내역 Matrix의 필드 Hidden
'    oMat06.Columns("CpCode").Visible = False '공정코드

    '자체가공비내역의 공정코드 필드 Hidden
    oForm01.Items("CpCode").Visible = False
    
    Call oForm01.Freeze(False)
    Exit Function
PS_PP362_CreateItems_Error:
    
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP362_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_PP362_ComboBox_Setting()
On Error GoTo PS_PP362_ComboBox_Setting_Error
    
    Call oForm01.Freeze(True)

    '거래처구분_S
    Call oForm01.Items("CardType").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("CardType").Specific, "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'C100' ORDER BY Code", "", False, False)
    Call oForm01.Items("CardType").Specific.Select(0, psk_Index)
    '거래처구분_E
    
    '품목구분_S
    Call oForm01.Items("ItemClass").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("ItemClass").Specific, "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'S002' ORDER BY Code", "", False, False)
    Call oForm01.Items("ItemClass").Specific.Select(0, psk_Index)
    '품목구분_E
    
    '생산완료여부_S
    Call oForm01.Items("WCYN").Specific.ValidValues.Add("%", "전체")
    Call oForm01.Items("WCYN").Specific.ValidValues.Add("B", "미완료")
    Call oForm01.Items("WCYN").Specific.ValidValues.Add("C", "완료")
    Call oForm01.Items("WCYN").Specific.Select(0, psk_Index)
    '생산완료여부_E

    Call oForm01.Freeze(False)
    
    Exit Sub
PS_PP362_ComboBox_Setting_Error:
'    Set lRecordSet = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP362_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP362_Initial_Setting()
On Error GoTo PS_PP362_Initial_Setting_Error

    '사업장 사용자의 소속 사업장 선택
'    Call oForm01.Items("BPLId").Specific.Select(MDC_PS_Common.User_BPLId(), psk_ByValue)
    
    '기준일자
'    oForm01.Items("StdDt").Specific.VALUE = Format(Now, "YYYYMMDD")
    

    
    '날짜 설정
'    oForm01.Items("ToDt").Specific.VALUE = Format(Date, "YYYYMMDD")
'    oForm01.Items("FrDt").Specific.VALUE = Format(Now, "YYYYMM") + "01"
    '날짜 빈칸으로 설정
'    oForm01.Items("ToDt").Specific.VALUE = ""
'    oForm01.Items("FrDt").Specific.VALUE = ""
 
    Exit Sub
PS_PP362_Initial_Setting_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP362_Initial_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Sub PS_PP362_CF_ChooseFromList()
On Error GoTo PS_PP362_CF_ChooseFromList_Error
    Exit Sub
PS_PP362_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP362_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP362_FormItemEnabled()
On Error GoTo PS_PP362_FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then

    ElseIf (oForm01.Mode = fm_FIND_MODE) Then

    ElseIf (oForm01.Mode = fm_OK_MODE) Then

    End If
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP362_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP362_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP362_AddMatrixRow(ByVal oRow As Long, ByVal prmMat As SAPbouiCOM.Matrix, ByVal prmDataSource As SAPbouiCOM.DBDataSource, Optional RowIserted As Boolean)
On Error GoTo PS_PP362_AddMatrixRow_Error
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//행추가여부
        prmDataSource.InsertRecord (oRow)
    End If
    prmMat.AddRow
    prmDataSource.Offset = oRow
'    oDS_PS_PP362L.setValue "U_LineNum", oRow, oRow + 1
    prmMat.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_PP362_AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP362_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_PP362_FormClear()
On Error GoTo PS_PP362_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_PP362'", "")
    If DocEntry = 0 Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PS_PP362_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP362_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP362_EnableMenus()
On Error GoTo PS_PP362_EnableMenus_Error
    Exit Sub
PS_PP362_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP362_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP362_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_PP362_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_PP362_FormItemEnabled

    Else

    End If
    Exit Sub
PS_PP362_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP362_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Function PS_PP362_DataValidCheck() As Boolean
On Error GoTo PS_PP362_DataValidCheck_Error
    PS_PP362_DataValidCheck = False
    Dim i As Long

'    If oForm01.Items("DocDateTo").Specific.Value = "" Then
'        Sbo_Application.SetStatusBarMessage "견적일자 To은 필수입니다.", bmt_Short, True
'        oForm01.Items("DocDateTo").Click ct_Regular
'        PS_PP362_DataValidCheck = False
'        Exit Function
'    End If

    PS_PP362_DataValidCheck = True
    Exit Function
PS_PP362_DataValidCheck_Error:
    PS_PP362_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PS_PP362_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_PP362_MTX01()
'******************************************************************************
'Function ID : PS_PP362_MTX01()
'해당모듈 : PS_PP362
'기능 : Main작번 TAB 내용 조회
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_PP362_MTX01_Error
    
    Dim loopCount As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim FrDt As String '작번등록년월(시작)
    Dim ToDt As String '작번등록년월(종료)
    Dim FrgnName As String '품명
    Dim CardType As String '거래처구분
    Dim Spec As String '규격
    Dim ItemClass As String '품목구분
    Dim WCYN As String '생산완료여부
    Dim CardCode As String '거래처
    Dim ItemCode As String '품목(작번)
    
    FrDt = oForm01.Items("FrDt").Specific.VALUE
    ToDt = oForm01.Items("ToDt").Specific.VALUE
    FrgnName = IIf(oForm01.Items("FrgnName").Specific.VALUE = "", "%", oForm01.Items("FrgnName").Specific.VALUE)
    CardType = oForm01.Items("CardType").Specific.Selected.VALUE
    Spec = IIf(oForm01.Items("Spec").Specific.VALUE = "", "%", oForm01.Items("Spec").Specific.VALUE)
    ItemClass = oForm01.Items("ItemClass").Specific.Selected.VALUE
    WCYN = IIf(oForm01.Items("WCYN").Specific.Selected.VALUE = "", "%", oForm01.Items("WCYN").Specific.Selected.VALUE)
    CardCode = IIf(oForm01.Items("CardCode").Specific.VALUE = "", "%", oForm01.Items("CardCode").Specific.VALUE)
    ItemCode = IIf(oForm01.Items("ItemCode").Specific.VALUE = "", "%", oForm01.Items("ItemCode").Specific.VALUE)
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, False) '쿼리를 실행할 때 부터 프로그레스 시작
    
    Call oForm01.Freeze(True)
    
    Query01 = "EXEC PS_PP362_01 '" & FrDt & "','" & ToDt & "','" & FrgnName & "','" & CardType & "','" & Spec & "','" & ItemClass & "','" & WCYN & "','" & CardCode & "','" & ItemCode & "'"
    Call RecordSet01.DoQuery(Query01)
    
    oMat01.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource
    
    If (RecordSet01.RecordCount = 0) Then
        Call oMat01.Clear
        GoTo PS_PP362_MTX01_Exit
    End If
    
    For loopCount = 0 To RecordSet01.RecordCount - 1
        If loopCount <> 0 Then
           Call oDS_PS_PP362L.InsertRecord(loopCount)
        End If
        oDS_PS_PP362L.Offset = loopCount
        
        Call oDS_PS_PP362L.setValue("U_LineNum", loopCount, loopCount + 1) '라인번호
        Call oDS_PS_PP362L.setValue("U_ColReg01", loopCount, RecordSet01.Fields("OrdNum").VALUE) '작번
        Call oDS_PS_PP362L.setValue("U_ColReg02", loopCount, RecordSet01.Fields("CardName").VALUE) '납품처
        Call oDS_PS_PP362L.setValue("U_ColReg03", loopCount, RecordSet01.Fields("ItemCode").VALUE) '품목
        Call oDS_PS_PP362L.setValue("U_ColReg04", loopCount, RecordSet01.Fields("ItemName").VALUE) '품목명
        Call oDS_PS_PP362L.setValue("U_ColReg05", loopCount, RecordSet01.Fields("Spec").VALUE) '규격
        Call oDS_PS_PP362L.setValue("U_ColReg06", loopCount, RecordSet01.Fields("Unit").VALUE) '단위
        Call oDS_PS_PP362L.setValue("U_ColSum01", loopCount, RecordSet01.Fields("Quantity").VALUE) '수량
        Call oDS_PS_PP362L.setValue("U_ColSum02", loopCount, RecordSet01.Fields("MatAmt").VALUE) '자재비
        Call oDS_PS_PP362L.setValue("U_ColSum03", loopCount, RecordSet01.Fields("DrawAmt").VALUE) '설계비
        Call oDS_PS_PP362L.setValue("U_ColSum04", loopCount, RecordSet01.Fields("GagongAmt").VALUE) '자체가공비
        Call oDS_PS_PP362L.setValue("U_ColSum05", loopCount, RecordSet01.Fields("OutgAmt").VALUE) '외주가공비
        Call oDS_PS_PP362L.setValue("U_ColSum06", loopCount, RecordSet01.Fields("OutmAmt").VALUE) '외주제작비
        Call oDS_PS_PP362L.setValue("U_ColSum07", loopCount, RecordSet01.Fields("Total").VALUE) '계
        Call oDS_PS_PP362L.setValue("U_ColSum08", loopCount, RecordSet01.Fields("SjAmt").VALUE) '수주금액
        Call oDS_PS_PP362L.setValue("U_ColReg15", loopCount, RecordSet01.Fields("DocDate").VALUE) '수주일자
        Call oDS_PS_PP362L.setValue("U_ColReg16", loopCount, RecordSet01.Fields("ShipDate").VALUE) '납기일자
        Call oDS_PS_PP362L.setValue("U_ColReg17", loopCount, RecordSet01.Fields("EndDate").VALUE) '완료일자
        Call oDS_PS_PP362L.setValue("U_ColReg18", loopCount, RecordSet01.Fields("YesNo").VALUE) '완료구분
        
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
    
PS_PP362_MTX01_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "W")
    Exit Sub
PS_PP362_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP362_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP362_MTX02(ByVal prmOrdNum As String)
'******************************************************************************
'Function ID : PS_PP362_MTX02()
'해당모듈 : PS_PP362
'기능 : Sub작번 TAB 내용 조회
'인수 : prmOrdNum(작번)
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_PP362_MTX02_Error
    
    Dim loopCount As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, False) '쿼리를 실행할 때 부터 프로그레스 시작
    
    Call oForm01.Freeze(True)
    
    Query01 = "EXEC PS_PP362_02 '" & prmOrdNum & "'"
    Call RecordSet01.DoQuery(Query01)
    
    oMat02.Clear
    oMat02.FlushToDataSource
    oMat02.LoadFromDataSource
    
    If (RecordSet01.RecordCount = 0) Then
        Call oMat02.Clear
        GoTo PS_PP362_MTX02_Exit
    End If
    
    For loopCount = 0 To RecordSet01.RecordCount - 1
        If loopCount <> 0 Then
           Call oDS_PS_PP362M.InsertRecord(loopCount)
        End If
        oDS_PS_PP362M.Offset = loopCount
        
        Call oDS_PS_PP362M.setValue("U_LineNum", loopCount, loopCount + 1) '라인번호
        Call oDS_PS_PP362M.setValue("U_ColReg01", loopCount, RecordSet01.Fields("OrdNum").VALUE) '작번
        Call oDS_PS_PP362M.setValue("U_ColReg02", loopCount, RecordSet01.Fields("OrdSub").VALUE) 'Sub작번
        Call oDS_PS_PP362M.setValue("U_ColReg03", loopCount, RecordSet01.Fields("ItemCode").VALUE) '품목
        Call oDS_PS_PP362M.setValue("U_ColReg04", loopCount, RecordSet01.Fields("ItemName").VALUE) '품목명
        Call oDS_PS_PP362M.setValue("U_ColReg05", loopCount, RecordSet01.Fields("Spec").VALUE) '규격
        Call oDS_PS_PP362M.setValue("U_ColReg06", loopCount, RecordSet01.Fields("Unit").VALUE) '단위
        Call oDS_PS_PP362M.setValue("U_ColSum01", loopCount, RecordSet01.Fields("Quantity").VALUE) '수량
        Call oDS_PS_PP362M.setValue("U_ColSum02", loopCount, RecordSet01.Fields("MatAmt").VALUE) '자재비
        Call oDS_PS_PP362M.setValue("U_ColSum03", loopCount, RecordSet01.Fields("DrawAmt").VALUE) '설계비
        Call oDS_PS_PP362M.setValue("U_ColSum04", loopCount, RecordSet01.Fields("GagongAmt").VALUE) '자체가공비
        Call oDS_PS_PP362M.setValue("U_ColSum05", loopCount, RecordSet01.Fields("OutgAmt").VALUE) '외주가공비
        Call oDS_PS_PP362M.setValue("U_ColSum06", loopCount, RecordSet01.Fields("OutmAmt").VALUE) '외주제작비
        Call oDS_PS_PP362M.setValue("U_ColSum07", loopCount, RecordSet01.Fields("Total").VALUE) '계
        Call oDS_PS_PP362M.setValue("U_ColSum08", loopCount, RecordSet01.Fields("SjAmt").VALUE) '수주금액
        Call oDS_PS_PP362M.setValue("U_ColReg15", loopCount, RecordSet01.Fields("PP030HNo").VALUE) '작지등록No
        
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
    
PS_PP362_MTX02_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "W")
    Exit Sub
PS_PP362_MTX02_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP362_MTX02_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP362_MTX03(ByVal prmOrdNum As String, ByVal prmOrdSub As String)
'******************************************************************************
'Function ID : PS_PP362_MTX03()
'해당모듈 : PS_PP362
'기능 : 자재비내역 TAB 내용 조회
'인수 : prmOrdNum(작번), prmOrdSub(Sub작번)
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_PP362_MTX03_Error
    
    Dim loopCount As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, False) '쿼리를 실행할 때 부터 프로그레스 시작
    
    Call oForm01.Freeze(True)
    
    Query01 = "EXEC PS_PP362_03 '" & prmOrdNum & "','" & prmOrdSub & "'"
    Call RecordSet01.DoQuery(Query01)
    
    oMat03.Clear
    oMat03.FlushToDataSource
    oMat03.LoadFromDataSource
    
    If (RecordSet01.RecordCount = 0) Then
        Call oMat03.Clear
        GoTo PS_PP362_MTX03_Exit
    End If
    
    For loopCount = 0 To RecordSet01.RecordCount - 1
        If loopCount <> 0 Then
           Call oDS_PS_PP362N.InsertRecord(loopCount)
        End If
        oDS_PS_PP362N.Offset = loopCount
        
        Call oDS_PS_PP362N.setValue("U_LineNum", loopCount, loopCount + 1) '라인번호
        Call oDS_PS_PP362N.setValue("U_ColReg01", loopCount, RecordSet01.Fields("Purchase").VALUE) '구분
        Call oDS_PS_PP362N.setValue("U_ColReg02", loopCount, RecordSet01.Fields("PurchaseNm").VALUE) '구분명
        Call oDS_PS_PP362N.setValue("U_ColReg03", loopCount, RecordSet01.Fields("OrdNum").VALUE) '작번
        Call oDS_PS_PP362N.setValue("U_ColReg04", loopCount, RecordSet01.Fields("ItemCode").VALUE) '품목코드
        Call oDS_PS_PP362N.setValue("U_ColReg05", loopCount, RecordSet01.Fields("ItemName").VALUE) '품명
        Call oDS_PS_PP362N.setValue("U_ColSum01", loopCount, RecordSet01.Fields("MatAmt").VALUE) '자재비
        
        RecordSet01.MoveNext
        ProgressBar01.VALUE = ProgressBar01.VALUE + 1
        ProgressBar01.Text = ProgressBar01.VALUE & "/" & RecordSet01.RecordCount & "건 조회중...!"
    Next
    oMat03.LoadFromDataSource
    oMat03.AutoResizeColumns
    
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
    
PS_PP362_MTX03_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
'    Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "W")
    Exit Sub
PS_PP362_MTX03_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP362_MTX03_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP362_MTX04(ByVal prmOrdNum As String, ByVal prmOrdSub As String)
'******************************************************************************
'Function ID : PS_PP362_MTX04()
'해당모듈 : PS_PP362
'기능 : 설계비내역 TAB 내용 조회
'인수 : prmOrdNum(작번), prmOrdSub(Sub작번)
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_PP362_MTX04_Error
    
    Dim loopCount As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, False) '쿼리를 실행할 때 부터 프로그레스 시작
    
    Call oForm01.Freeze(True)
    
    Query01 = "EXEC PS_PP362_04 '" & prmOrdNum & "','" & prmOrdSub & "'"
    Call RecordSet01.DoQuery(Query01)
    
    oMat04.Clear
    oMat04.FlushToDataSource
    oMat04.LoadFromDataSource
    
    If (RecordSet01.RecordCount = 0) Then
        Call oMat04.Clear
        GoTo PS_PP362_MTX04_Exit
    End If
    
    For loopCount = 0 To RecordSet01.RecordCount - 1
        If loopCount <> 0 Then
           Call oDS_PS_PP362O.InsertRecord(loopCount)
        End If
        oDS_PS_PP362O.Offset = loopCount
        
        Call oDS_PS_PP362O.setValue("U_LineNum", loopCount, loopCount + 1) '라인번호
        Call oDS_PS_PP362O.setValue("U_ColReg01", loopCount, RecordSet01.Fields("DocDate").VALUE) '일자
        Call oDS_PS_PP362O.setValue("U_ColReg02", loopCount, RecordSet01.Fields("WorkCode").VALUE) '작업자코드
        Call oDS_PS_PP362O.setValue("U_ColReg03", loopCount, RecordSet01.Fields("WorkName").VALUE) '작업자명
        Call oDS_PS_PP362O.setValue("U_ColReg04", loopCount, RecordSet01.Fields("CpCode").VALUE) '공정코드
        Call oDS_PS_PP362O.setValue("U_ColReg05", loopCount, RecordSet01.Fields("CpName").VALUE) '공정명
        Call oDS_PS_PP362O.setValue("U_ColReg06", loopCount, RecordSet01.Fields("PQty").VALUE) '도면매수
        Call oDS_PS_PP362O.setValue("U_ColSum01", loopCount, RecordSet01.Fields("Amt").VALUE) '설계비
        
        RecordSet01.MoveNext
        ProgressBar01.VALUE = ProgressBar01.VALUE + 1
        ProgressBar01.Text = ProgressBar01.VALUE & "/" & RecordSet01.RecordCount & "건 조회중...!"
    Next
    oMat04.LoadFromDataSource
    oMat04.AutoResizeColumns
    
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
    
PS_PP362_MTX04_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
'    Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "W")
    Exit Sub
PS_PP362_MTX04_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP362_MTX04_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP362_MTX05(ByVal prmOrdNum As String, ByVal prmOrdSub As String)
'******************************************************************************
'Function ID : PS_PP362_MTX05()
'해당모듈 : PS_PP362
'기능 : 자체가공비내역(공정별내역) TAB 내용 조회
'인수 : prmOrdNum(작번), prmOrdSub(Sub작번)
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_PP362_MTX05_Error
    
    Dim loopCount As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, False) '쿼리를 실행할 때 부터 프로그레스 시작
    
    Call oForm01.Freeze(True)
    
    Query01 = "EXEC PS_PP362_05 '" & prmOrdNum & "','" & prmOrdSub & "'"
    Call RecordSet01.DoQuery(Query01)
    
    oMat05.Clear
    oMat05.FlushToDataSource
    oMat05.LoadFromDataSource
    
    If (RecordSet01.RecordCount = 0) Then
        Call oMat05.Clear
        GoTo PS_PP362_MTX05_Exit
    End If
    
    For loopCount = 0 To RecordSet01.RecordCount - 1
        If loopCount <> 0 Then
           Call oDS_PS_PP362P.InsertRecord(loopCount)
        End If
        oDS_PS_PP362P.Offset = loopCount
        
        Call oDS_PS_PP362P.setValue("U_LineNum", loopCount, loopCount + 1) '라인번호
        Call oDS_PS_PP362P.setValue("U_ColReg01", loopCount, RecordSet01.Fields("CpName").VALUE) '공정명
        Call oDS_PS_PP362P.setValue("U_ColReg02", loopCount, RecordSet01.Fields("StdTime").VALUE) '표준공수
        Call oDS_PS_PP362P.setValue("U_ColReg03", loopCount, RecordSet01.Fields("WkTime").VALUE) '실동공수
        Call oDS_PS_PP362P.setValue("U_ColReg04", loopCount, RecordSet01.Fields("ReTime").VALUE) '수정공수
        Call oDS_PS_PP362P.setValue("U_ColSum01", loopCount, RecordSet01.Fields("Amt").VALUE) '가공비(실동)
        Call oDS_PS_PP362P.setValue("U_ColReg06", loopCount, RecordSet01.Fields("Class").VALUE) '구분
        Call oDS_PS_PP362P.setValue("U_ColReg07", loopCount, RecordSet01.Fields("CompltDt").VALUE) '완료요구일
        Call oDS_PS_PP362P.setValue("U_ColReg08", loopCount, RecordSet01.Fields("FirstWkDt").VALUE) '최초작업
        Call oDS_PS_PP362P.setValue("U_ColReg09", loopCount, RecordSet01.Fields("LastWkDt").VALUE) '최종작업
        Call oDS_PS_PP362P.setValue("U_ColReg10", loopCount, RecordSet01.Fields("CpCode").VALUE) '공정코드
        Call oDS_PS_PP362P.setValue("U_ColReg11", loopCount, RecordSet01.Fields("OrdNum").VALUE) '작번
        Call oDS_PS_PP362P.setValue("U_ColReg12", loopCount, RecordSet01.Fields("Sub1_2").VALUE) '서브작번
        
        RecordSet01.MoveNext
        ProgressBar01.VALUE = ProgressBar01.VALUE + 1
        ProgressBar01.Text = ProgressBar01.VALUE & "/" & RecordSet01.RecordCount & "건 조회중...!"
    Next
    oMat05.LoadFromDataSource
    oMat05.AutoResizeColumns
    
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
    
PS_PP362_MTX05_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
'    Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "W")
    Exit Sub
PS_PP362_MTX05_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP362_MTX05_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP362_MTX06(ByVal prmOrdNum As String, ByVal prmOrdSub As String, ByVal prmCpCode As String)
'******************************************************************************
'Function ID : PS_PP362_MTX06()
'해당모듈 : PS_PP362
'기능 : 자체가공비내역(작업자별내역) TAB 내용 조회
'인수 : prmOrdNum(작번), prmOrdSub(Sub작번), prmCpCode(공정코드)
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_PP362_MTX06_Error
    
    Dim loopCount As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, False) '쿼리를 실행할 때 부터 프로그레스 시작
    
    Call oForm01.Freeze(True)
    
    Query01 = "EXEC PS_PP362_06 '" & prmOrdNum & "','" & prmOrdSub & "','" & prmCpCode & "'"
    Call RecordSet01.DoQuery(Query01)
    
    oMat06.Clear
    oMat06.FlushToDataSource
    oMat06.LoadFromDataSource
    
    If (RecordSet01.RecordCount = 0) Then
        Call oMat06.Clear
        GoTo PS_PP362_MTX06_Exit
    End If
    
    For loopCount = 0 To RecordSet01.RecordCount - 1
        If loopCount <> 0 Then
           Call oDS_PS_PP362Q.InsertRecord(loopCount)
        End If
        oDS_PS_PP362Q.Offset = loopCount
        
        Call oDS_PS_PP362Q.setValue("U_LineNum", loopCount, loopCount + 1) '라인번호
        Call oDS_PS_PP362Q.setValue("U_ColReg01", loopCount, RecordSet01.Fields("WorkDt").VALUE) '작업일자
        Call oDS_PS_PP362Q.setValue("U_ColReg02", loopCount, RecordSet01.Fields("EmpCode").VALUE) '사번
        Call oDS_PS_PP362Q.setValue("U_ColReg03", loopCount, RecordSet01.Fields("EmpName").VALUE) '성명
        Call oDS_PS_PP362Q.setValue("U_ColQty01", loopCount, RecordSet01.Fields("WkTime").VALUE) '공수
        Call oDS_PS_PP362Q.setValue("U_ColReg05", loopCount, RecordSet01.Fields("Qty").VALUE) '완료수량
        Call oDS_PS_PP362Q.setValue("U_ColReg06", loopCount, RecordSet01.Fields("Price").VALUE) '단가
        Call oDS_PS_PP362Q.setValue("U_ColSum02", loopCount, RecordSet01.Fields("Amt").VALUE) '금액(가공비)
        Call oDS_PS_PP362Q.setValue("U_ColReg08", loopCount, RecordSet01.Fields("Class").VALUE) '구분
        Call oDS_PS_PP362Q.setValue("U_ColReg09", loopCount, RecordSet01.Fields("CpCode").VALUE) '공정코드
        Call oDS_PS_PP362Q.setValue("U_ColReg10", loopCount, RecordSet01.Fields("CpName").VALUE) '공정명
        Call oDS_PS_PP362Q.setValue("U_ColReg11", loopCount, RecordSet01.Fields("PP040HNo").VALUE) '작업일보번호
        
        RecordSet01.MoveNext
        ProgressBar01.VALUE = ProgressBar01.VALUE + 1
        ProgressBar01.Text = ProgressBar01.VALUE & "/" & RecordSet01.RecordCount & "건 조회중...!"
    Next
    oMat06.LoadFromDataSource
    oMat06.AutoResizeColumns
    
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
    
PS_PP362_MTX06_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
'    Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "W")
    Exit Sub
PS_PP362_MTX06_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP362_MTX06_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP362_MTX07(ByVal prmOrdNum As String, ByVal prmOrdSub As String)
'******************************************************************************
'Function ID : PS_PP362_MTX07()
'해당모듈 : PS_PP362
'기능 : 외주가공비내역 TAB 내용 조회
'인수 : prmOrdNum(작번), prmOrdSub(Sub작번)
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_PP362_MTX07_Error
    
    Dim loopCount As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, False) '쿼리를 실행할 때 부터 프로그레스 시작
    
    Call oForm01.Freeze(True)
    
    Query01 = "EXEC PS_PP362_07 '" & prmOrdNum & "','" & prmOrdSub & "'"
    Call RecordSet01.DoQuery(Query01)
    
    oMat07.Clear
    oMat07.FlushToDataSource
    oMat07.LoadFromDataSource
    
    If (RecordSet01.RecordCount = 0) Then
        Call oMat07.Clear
        GoTo PS_PP362_MTX07_Exit
    End If
    
    For loopCount = 0 To RecordSet01.RecordCount - 1
        If loopCount <> 0 Then
           Call oDS_PS_PP362R.InsertRecord(loopCount)
        End If
        oDS_PS_PP362R.Offset = loopCount
        
        Call oDS_PS_PP362R.setValue("U_LineNum", loopCount, loopCount + 1) '라인번호
        Call oDS_PS_PP362R.setValue("U_ColReg01", loopCount, RecordSet01.Fields("CpCode").VALUE) '공정코드
        Call oDS_PS_PP362R.setValue("U_ColReg02", loopCount, RecordSet01.Fields("CpName").VALUE) '공정명
        Call oDS_PS_PP362R.setValue("U_ColReg03", loopCount, RecordSet01.Fields("OrdNum").VALUE) '작번
        Call oDS_PS_PP362R.setValue("U_ColReg04", loopCount, RecordSet01.Fields("ItemName").VALUE) '품명
        Call oDS_PS_PP362R.setValue("U_ColSum01", loopCount, RecordSet01.Fields("Amt").VALUE) '금액
        
        RecordSet01.MoveNext
        ProgressBar01.VALUE = ProgressBar01.VALUE + 1
        ProgressBar01.Text = ProgressBar01.VALUE & "/" & RecordSet01.RecordCount & "건 조회중...!"
    Next
    oMat07.LoadFromDataSource
    oMat07.AutoResizeColumns
    
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
    
PS_PP362_MTX07_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
'    Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "W")
    Exit Sub
PS_PP362_MTX07_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP362_MTX07_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP362_MTX08(ByVal prmOrdNum As String, ByVal prmOrdSub As String)
'******************************************************************************
'Function ID : PS_PP362_MTX08()
'해당모듈 : PS_PP362
'기능 : 외주제작비내역 TAB 내용 조회
'인수 : prmOrdNum(작번), prmOrdSub(Sub작번)
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_PP362_MTX08_Error
    
    Dim loopCount As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, False) '쿼리를 실행할 때 부터 프로그레스 시작
    
    Call oForm01.Freeze(True)
    
    Query01 = "EXEC PS_PP362_08 '" & prmOrdNum & "','" & prmOrdSub & "'"
    Call RecordSet01.DoQuery(Query01)
    
    oMat08.Clear
    oMat08.FlushToDataSource
    oMat08.LoadFromDataSource
    
    If (RecordSet01.RecordCount = 0) Then
        Call oMat08.Clear
        GoTo PS_PP362_MTX08_Exit
    End If
    
    For loopCount = 0 To RecordSet01.RecordCount - 1
        If loopCount <> 0 Then
           Call oDS_PS_PP362S.InsertRecord(loopCount)
        End If
        oDS_PS_PP362S.Offset = loopCount
        
        Call oDS_PS_PP362S.setValue("U_LineNum", loopCount, loopCount + 1) '라인번호
        Call oDS_PS_PP362S.setValue("U_ColReg01", loopCount, RecordSet01.Fields("CpCode").VALUE) '공정코드
        Call oDS_PS_PP362S.setValue("U_ColReg02", loopCount, RecordSet01.Fields("CpName").VALUE) '공정명
        Call oDS_PS_PP362S.setValue("U_ColReg03", loopCount, RecordSet01.Fields("OrdNum").VALUE) '작번
        Call oDS_PS_PP362S.setValue("U_ColReg04", loopCount, RecordSet01.Fields("ItemName").VALUE) '품명
        Call oDS_PS_PP362S.setValue("U_ColSum01", loopCount, RecordSet01.Fields("Amt").VALUE) '금액
        
        RecordSet01.MoveNext
        ProgressBar01.VALUE = ProgressBar01.VALUE + 1
        ProgressBar01.Text = ProgressBar01.VALUE & "/" & RecordSet01.RecordCount & "건 조회중...!"
    Next
    oMat08.LoadFromDataSource
    oMat08.AutoResizeColumns
    
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
    
PS_PP362_MTX08_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
'    Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "W")
    Exit Sub
PS_PP362_MTX08_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_PP362_MTX08_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP362_FormResize()
On Error GoTo PS_PP362_FormResize_Error
    
    oForm01.Items("Mat05").Width = oForm01.Width / 2 - 25
    oForm01.Items("Mat05").Height = oForm01.Height - oForm01.Items("Mat05").Top - (oForm01.Height - oForm01.Items("Btn05").Top) - 10

    oForm01.Items("Mat06").Left = oForm01.Width / 2 + 5
    oForm01.Items("Mat06").Width = oForm01.Width / 2 - 45
    oForm01.Items("Mat06").Height = oForm01.Items("Mat05").Height

    oForm01.Items("Static11").Left = oForm01.Items("Mat06").Left
    
    oForm01.Items("Btn06").Top = oForm01.Items("Btn05").Top
    oForm01.Items("Btn06").Left = oForm01.Items("Static11").Left
        
    oForm01.Items("Rec01").Height = oForm01.Items("Mat01").Height + 50
    oForm01.Items("Rec01").Width = oForm01.Items("Mat01").Width + 25
    
    oMat01.AutoResizeColumns
    oMat02.AutoResizeColumns
    oMat03.AutoResizeColumns
    oMat04.AutoResizeColumns
    oMat05.AutoResizeColumns
    oMat06.AutoResizeColumns
    oMat07.AutoResizeColumns
    oMat08.AutoResizeColumns
    
    Exit Sub
PS_PP362_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_PP362_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_PP362_Print_Report(ByVal prmItemUID As String)
On Error GoTo PS_PP362_Print_Report_Error

    Dim loopCount As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
   
    Call ConnectODBC
   
    Dim ErrNum As Integer
    Dim WinTitle As String
    Dim ReportName As String
    
    WinTitle = "[PS_PP362] 레포트"
    
    Dim FrDt As String '작번등록년월(시작)
    Dim ToDt As String '작번등록년월(종료)
    Dim FrgnName As String '품명
    Dim CardType As String '거래처구분
    Dim Spec As String '규격
    Dim ItemClass As String '품목구분
    Dim WCYN As String '생산완료여부
    Dim CardCode As String '거래처
    Dim ItemCode As String '품목(작번)
    
    Dim OrdNum As String 'Main작번
    Dim OrdSub As String 'Sub작번
    Dim CpCode As String '공정코드
    
    If prmItemUID = "Btn01" Then 'Main작번 출력
    
        FrDt = oForm01.Items("FrDt").Specific.VALUE
        ToDt = oForm01.Items("ToDt").Specific.VALUE
        FrgnName = IIf(oForm01.Items("FrgnName").Specific.VALUE = "", "%", oForm01.Items("FrgnName").Specific.VALUE)
        CardType = oForm01.Items("CardType").Specific.Selected.VALUE
        Spec = IIf(oForm01.Items("Spec").Specific.VALUE = "", "%", oForm01.Items("Spec").Specific.VALUE)
        ItemClass = oForm01.Items("ItemClass").Specific.Selected.VALUE
        WCYN = IIf(oForm01.Items("WCYN").Specific.Selected.VALUE = "", "%", oForm01.Items("WCYN").Specific.Selected.VALUE)
        CardCode = IIf(oForm01.Items("CardCode").Specific.VALUE = "", "%", oForm01.Items("CardCode").Specific.VALUE)
        ItemCode = IIf(oForm01.Items("ItemCode").Specific.VALUE = "", "%", oForm01.Items("ItemCode").Specific.VALUE)
       
        ReportName = "PS_PP362_01.rpt"
    
        Query01 = "EXEC PS_PP362_01 '" & FrDt & "','" & ToDt & "','" & FrgnName & "','" & CardType & "','" & Spec & "','" & ItemClass & "','" & WCYN & "','" & CardCode & "','" & ItemCode & "'"
       
    ElseIf prmItemUID = "Btn02" Then 'Sub작번 출력
    
        OrdNum = oMat02.Columns("OrdNum").Cells(1).Specific.VALUE
    
        ReportName = "PS_PP362_02.rpt"
    
        Query01 = "EXEC PS_PP362_02 '" & OrdNum & "'"
    
    ElseIf prmItemUID = "Btn03" Then '자재비내역 출력
    
        OrdNum = oForm01.Items("MainJak").Specific.VALUE
        OrdSub = oForm01.Items("SubJak").Specific.VALUE
        
        ReportName = "PS_PP362_03.rpt"
    
        Query01 = "EXEC PS_PP362_03 '" & OrdNum & "','" & OrdSub & "'"
        
    ElseIf prmItemUID = "Btn04" Then '설계비내역 출력
    
        OrdNum = oForm01.Items("MainJak").Specific.VALUE
        OrdSub = oForm01.Items("SubJak").Specific.VALUE
        
        ReportName = "PS_PP362_04.rpt"
    
        Query01 = "EXEC PS_PP362_04 '" & OrdNum & "','" & OrdSub & "'"
    
    ElseIf prmItemUID = "Btn05" Then '자체가공비내역(공정별) 출력
    
        OrdNum = oForm01.Items("MainJak").Specific.VALUE
        OrdSub = oForm01.Items("SubJak").Specific.VALUE
        
        ReportName = "PS_PP362_05.rpt"
    
        Query01 = "EXEC PS_PP362_05 '" & OrdNum & "','" & OrdSub & "'"
    
    ElseIf prmItemUID = "Btn06" Then '자체가공비내역(작업자별) 출력
    
        OrdNum = oForm01.Items("MainJak").Specific.VALUE
        OrdSub = oForm01.Items("SubJak").Specific.VALUE
        CpCode = oForm01.Items("CpCode").Specific.VALUE
        
        ReportName = "PS_PP362_06.rpt"
    
        Query01 = "EXEC PS_PP362_06 '" & OrdNum & "','" & OrdSub & "','" & CpCode & "'"
        
    ElseIf prmItemUID = "Btn07" Then '외주가공비내역 출력
    
        OrdNum = oForm01.Items("MainJak").Specific.VALUE
        OrdSub = oForm01.Items("SubJak").Specific.VALUE
        
        ReportName = "PS_PP362_07.rpt"
    
        Query01 = "EXEC PS_PP362_07 '" & OrdNum & "','" & OrdSub & "'"
    
    ElseIf prmItemUID = "Btn08" Then '외주제작비내역 출력
    
        OrdNum = oForm01.Items("MainJak").Specific.VALUE
        OrdSub = oForm01.Items("SubJak").Specific.VALUE
        
        ReportName = "PS_PP362_08.rpt"
    
        Query01 = "EXEC PS_PP362_08 '" & OrdNum & "','" & OrdSub & "'"
    
    End If


    '// Formula 수식필드
    ReDim gRpt_Formula(0)
    ReDim gRpt_Formula_Value(0)
       
    '// SubReport
    ReDim gRpt_SRptSqry(0)
    ReDim gRpt_SRptName(0)
    
    ReDim gRpt_SFormula(0, 0)
    ReDim gRpt_SFormula_Value(0, 0)
    
    gRpt_SFormula(0, 0) = ""
    gRpt_SFormula_Value(0, 0) = ""
    
    Call RecordSet01.DoQuery(Query01)

    If RecordSet01.RecordCount = 0 Then
        ErrNum = 1
        GoTo PS_PP362_Print_Report_Error
    End If

    '/ Action (sub_query가 있을때는 'Y'로...)/
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", Query01, "", "N", "V") = False Then
    End If
    
    Set RecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_PP362_Print_Report_Error:
    Set RecordSet01 = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "PS_PP362_Print_Report:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub




