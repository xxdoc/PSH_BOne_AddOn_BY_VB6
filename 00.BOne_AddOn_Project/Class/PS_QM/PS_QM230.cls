VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_QM230"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File : PS_QM230.cls
'//  Module : 품질관리 > 엔드베어링 품질경향 관리
'//  Desc : 엔드베어링 품질경향 등록
'//  FormType : PS_QM230
'//  Create Date(Start) : 2012.05.22
'//  Create Date(End) :
'//  Modified Date :
'//  Creator : Song Myoung gyu
'//  Company : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form

Public oMat01 As SAPbouiCOM.Matrix '외경Max
Public oMat02 As SAPbouiCOM.Matrix '외경Min
Public oMat03 As SAPbouiCOM.Matrix '내경Max
Public oMat04 As SAPbouiCOM.Matrix '내경Min
Public oMat05 As SAPbouiCOM.Matrix '높이Max
Public oMat06 As SAPbouiCOM.Matrix '높이Min

Private oDS_PS_QM230H As SAPbouiCOM.DBDataSource '헤더
Private oDS_PS_QM230L As SAPbouiCOM.DBDataSource '라인(외경Max)
Private oDS_PS_QM230M As SAPbouiCOM.DBDataSource '라인(외경Min)
Private oDS_PS_QM230N As SAPbouiCOM.DBDataSource '라인(내경Max)
Private oDS_PS_QM230O As SAPbouiCOM.DBDataSource '라인(내경Min)
Private oDS_PS_QM230P As SAPbouiCOM.DBDataSource '라인(높이Max)
Private oDS_PS_QM230Q As SAPbouiCOM.DBDataSource '라인(높이Min)

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
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_QM230.srf")
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
    
    oFormUniqueID01 = "PS_QM230_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    oForm01.DataBrowser.BrowseBy = "DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_QM230_CreateItems
    Call PS_QM230_ComboBox_Setting
    Call PS_QM230_CF_ChooseFromList
    Call PS_QM230_EnableMenus
    Call PS_QM230_SetDocument(oFromDocEntry01)
    Call PS_QM230_FormResize
    
    oForm01.EnableMenu ("1283"), False '삭제
    oForm01.EnableMenu ("1287"), False '복제
    oForm01.EnableMenu ("1286"), False '닫기
    oForm01.EnableMenu ("1284"), False '취소
    oForm01.EnableMenu ("1293"), True '행삭제
    oForm01.EnableMenu ("1299"), False '행닫기
    
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
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": '찾기
            Case "1282": '추가
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": '찾기
                Call PS_QM230_FormItemEnabled '//UDO방식
                oForm01.Items("DocEntry").Click ct_Regular
            Case "1282": '추가
                Call PS_QM230_FormItemEnabled '//UDO방식
                Call PS_QM230_AddMatrixRow01(0, True) '//UDO방식
                Call PS_QM230_AddMatrixRow02(0, True) '//UDO방식
                Call PS_QM230_AddMatrixRow03(0, True) '//UDO방식
                Call PS_QM230_AddMatrixRow04(0, True) '//UDO방식
                Call PS_QM230_AddMatrixRow05(0, True) '//UDO방식
                Call PS_QM230_AddMatrixRow06(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                Call PS_QM230_FormItemEnabled
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
    
        If pval.ItemUID = "PS_QM230" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        
        If pval.ItemUID = "1" Then
        
            If oForm01.Mode = fm_ADD_MODE Then
            
                '동일 년월, 동일 품목의 자료가 이미 입력되었는지 체크
                If PS_QM230_CheckDocument = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                
                If PS_QM230_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                '//해야할일 작업
                oDocEntry01 = Trim(oForm01.Items("DocEntry").Specific.VALUE)
                oFormMode01 = oForm01.Mode
                
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
                If PS_QM230_DataValidCheck = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                '//해야할일 작업
                oDocEntry01 = Trim(oForm01.Items("DocEntry").Specific.VALUE)
                oFormMode01 = oForm01.Mode
                
            ElseIf oForm01.Mode = fm_OK_MODE Then
            
            End If
            
        ElseIf pval.ItemUID = "BtnGraph01" Or pval.ItemUID = "BtnGraph02" Or pval.ItemUID = "BtnGraph03" Or pval.ItemUID = "BtnGraph04" Or pval.ItemUID = "BtnGraph05" Or pval.ItemUID = "BtnGraph06" Then
            
            Call PS_QM230_OpenGraph(oForm01.Items("StdYM").Specific.VALUE, oForm01.Items("ItemCode").Specific.VALUE, oForm01.PaneLevel)
            
        End If
        
    ElseIf pval.BeforeAction = False Then
    
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
        '폴더를 사용할 때는 필수 소스_E
    
        If pval.ItemUID = "PS_QM230" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        If pval.ItemUID = "1" Then
            If oForm01.Mode = fm_ADD_MODE Then
                If pval.ActionSuccess = True Then
                    Call PS_QM230_FormItemEnabled
                    Call PS_QM230_AddMatrixRow01(0, True) '//UDO방식일때
                    Call PS_QM230_AddMatrixRow02(0, True) '//UDO방식일때
                    Call PS_QM230_AddMatrixRow03(0, True) '//UDO방식일때
                    Call PS_QM230_AddMatrixRow04(0, True) '//UDO방식일때
                    Call PS_QM230_AddMatrixRow05(0, True) '//UDO방식일때
                    Call PS_QM230_AddMatrixRow06(0, True) '//UDO방식일때
                End If
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
                If pval.ActionSuccess = True Then
                    Call PS_QM230_FormItemEnabled
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
        
'            If pval.ColUID = "PGCd" Then
'
'                Call MDC_PS_Common.ActiveUserDefineValueAlways(oForm01, pval, BubbleEvent, "Mat01", "PGCd") '프로그램코드 포맷서치설정
'
'            End If
        
        ElseIf pval.ItemUID = "CntcCode" Then
        
            Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CntcCode", "") '등록자 포맷서치 설정
        
        ElseIf pval.ItemUID = "ItemCode" Then
        
            Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "ItemCode", "") '요청자 포맷서치 설정
        
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
    
'        If pval.ItemUID = "Mat01" Then
'
'            oMat01.FlushToDataSource
'
'            If pval.ColUID = "PGType" Then
'
'                Call oDS_PS_QM230L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.Selected.VALUE)
'                If oMat01.RowCount = pval.Row And Trim(oDS_PS_QM230L.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
'                    Call PS_QM230_AddMatrixRow01(pval.Row)
'                End If
'
'            End If
'
'        ElseIf pval.ItemUID = "Mat02" Then
'
'            oMat02.FlushToDataSource
'
'            If pval.ColUID = "PGType" Then
'
'                Call oDS_PS_QM230M.setValue("U_" & pval.ColUID, pval.Row - 1, oMat02.Columns(pval.ColUID).Cells(pval.Row).Specific.Selected.VALUE)
'                If oMat02.RowCount = pval.Row And Trim(oDS_PS_QM230M.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
'                    Call PS_QM230_AddMatrixRow02(pval.Row)
'                End If
'
'            End If
'
'        End If
        
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
            
        ElseIf pval.ItemUID = "Mat02" Then
        
            If pval.Row > 0 Then
                oLastItemUID01 = pval.ItemUID
                oLastColUID01 = pval.ColUID
                oLastColRow01 = pval.Row
                
                Call oMat02.SelectRow(pval.Row, True, False)
            End If
            
        ElseIf pval.ItemUID = "Mat03" Then
        
            If pval.Row > 0 Then
                oLastItemUID01 = pval.ItemUID
                oLastColUID01 = pval.ColUID
                oLastColRow01 = pval.Row
                
                Call oMat03.SelectRow(pval.Row, True, False)
            End If
            
        ElseIf pval.ItemUID = "Mat04" Then
        
            If pval.Row > 0 Then
                oLastItemUID01 = pval.ItemUID
                oLastColUID01 = pval.ColUID
                oLastColRow01 = pval.Row
                
                Call oMat04.SelectRow(pval.Row, True, False)
            End If
            
        ElseIf pval.ItemUID = "Mat05" Then
        
            If pval.Row > 0 Then
                oLastItemUID01 = pval.ItemUID
                oLastColUID01 = pval.ColUID
                oLastColRow01 = pval.Row
                
                Call oMat05.SelectRow(pval.Row, True, False)
            End If
            
        ElseIf pval.ItemUID = "Mat06" Then
        
            If pval.Row > 0 Then
                oLastItemUID01 = pval.ItemUID
                oLastColUID01 = pval.ColUID
                oLastColRow01 = pval.Row
                
                Call oMat06.SelectRow(pval.Row, True, False)
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
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim loopCount As Integer

    Call oForm01.Freeze(True)
    If pval.BeforeAction = True Then
        If pval.ItemChanged = True Then
        
            If pval.ItemUID = "Mat01" Then
            
                Call oMat01.AutoResizeColumns

            ElseIf pval.ItemUID = "Mat02" Then
            
                Call oMat02.AutoResizeColumns
            
            ElseIf pval.ItemUID = "Mat03" Then
            
                Call oMat03.AutoResizeColumns
            
            ElseIf pval.ItemUID = "Mat04" Then
            
                Call oMat04.AutoResizeColumns
            
            ElseIf pval.ItemUID = "Mat05" Then
            
                Call oMat05.AutoResizeColumns
            
            ElseIf pval.ItemUID = "Mat06" Then
            
                Call oMat06.AutoResizeColumns

            End If
        
            If pval.ItemUID = "CntcCode" Then '등록자
            
                Query01 = "SELECT U_FULLNAME FROM OHEM WHERE U_MSTCOD = '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'"
                Call RecordSet01.DoQuery(Query01)
                oForm01.Items("CntcName").Specific.VALUE = Trim(RecordSet01.Fields(0).VALUE)

            ElseIf pval.ItemUID = "ItemCode" Then '품목코드

                Query01 = "SELECT ItemName FROM OITM WHERE ItemCode = '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'"
                Call RecordSet01.DoQuery(Query01)
                oForm01.Items("ItemName").Specific.VALUE = Trim(RecordSet01.Fields(0).VALUE)
                
                If oMat01.RowCount = 1 Then
                
                    '품목코드를 선택하면 모든 매트릭스에 행 추가
                    For loopCount = 1 To 6
                    
                        Call PS_QM230_AddMatrixRow01(loopCount, False)
                        Call PS_QM230_AddMatrixRow02(loopCount, False)
                        Call PS_QM230_AddMatrixRow03(loopCount, False)
                        Call PS_QM230_AddMatrixRow04(loopCount, False)
                        Call PS_QM230_AddMatrixRow05(loopCount, False)
                        Call PS_QM230_AddMatrixRow06(loopCount, False)
                    
                    Next
                    
                    For loopCount = 1 To 6
                    
                        Call oMat01.Columns("cls").Cells(loopCount).Specific.Select(loopCount - 1, psk_Index)
                        Call oMat02.Columns("cls").Cells(loopCount).Specific.Select(loopCount - 1, psk_Index)
                        Call oMat03.Columns("cls").Cells(loopCount).Specific.Select(loopCount - 1, psk_Index)
                        Call oMat04.Columns("cls").Cells(loopCount).Specific.Select(loopCount - 1, psk_Index)
                        Call oMat05.Columns("cls").Cells(loopCount).Specific.Select(loopCount - 1, psk_Index)
                        Call oMat06.Columns("cls").Cells(loopCount).Specific.Select(loopCount - 1, psk_Index)
                    
                    Next
                    
                    Call oMat01.AutoResizeColumns
                    Call oMat02.AutoResizeColumns
                    Call oMat03.AutoResizeColumns
                    Call oMat04.AutoResizeColumns
                    Call oMat05.AutoResizeColumns
                    Call oMat06.AutoResizeColumns
                
                End If
                
            End If
        
            oForm01.Update
        
'            oForm01.Items(pval.ItemUID).Click ct_Regular
        
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
    
        Call PS_QM230_FormItemEnabled
        
        If pval.ItemUID = "Mat01" Then
        
            Call PS_QM230_AddMatrixRow01(oMat01.VisualRowCount) '//UDO방식
            Call oMat01.AutoResizeColumns
            
        ElseIf pval.ItemUID = "Mat02" Then
        
            Call PS_QM230_AddMatrixRow02(oMat02.VisualRowCount) '//UDO방식
            Call oMat02.AutoResizeColumns
            
        ElseIf pval.ItemUID = "Mat03" Then
        
            Call PS_QM230_AddMatrixRow03(oMat03.VisualRowCount) '//UDO방식
            Call oMat03.AutoResizeColumns
            
        ElseIf pval.ItemUID = "Mat04" Then
        
            Call PS_QM230_AddMatrixRow04(oMat04.VisualRowCount) '//UDO방식
            Call oMat04.AutoResizeColumns
            
        ElseIf pval.ItemUID = "Mat05" Then
        
            Call PS_QM230_AddMatrixRow05(oMat05.VisualRowCount) '//UDO방식
            Call oMat05.AutoResizeColumns
        
        ElseIf pval.ItemUID = "Mat06" Then
        
            Call PS_QM230_AddMatrixRow06(oMat06.VisualRowCount) '//UDO방식
            Call oMat06.AutoResizeColumns
        
        End If
        
    End If
    
    Exit Sub
    
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_QM230_FormResize
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
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_QM230H", "U_CardCode,U_CardName")
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
'            If (PS_QM230_Validate("행삭제") = False) Then
'                BubbleEvent = False
'                Exit Sub
'            End If
            '//행삭제전 행삭제가능여부검사
        ElseIf pval.BeforeAction = False Then
        
            If oLastItemUID01 = "Mat01" Then
        
                For i = 1 To oMat01.VisualRowCount
                    oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
                Next i
                
                oMat01.FlushToDataSource
                Call oDS_PS_QM230L.RemoveRecord(oDS_PS_QM230L.Size - 1)
                oMat01.LoadFromDataSource
                
                If oMat01.RowCount = 0 Then
                    Call PS_QM230_AddMatrixRow01(0)
                Else
                
                    If Trim(oDS_PS_QM230L.GetValue("U_cls", oMat01.RowCount - 1)) <> "" Then
                        Call PS_QM230_AddMatrixRow01(oMat01.RowCount)
                    End If
                    
                End If
                
            ElseIf oLastItemUID01 = "Mat02" Then
        
                For i = 1 To oMat02.VisualRowCount
                    oMat02.Columns("LineNum").Cells(i).Specific.VALUE = i
                Next i
                
                oMat02.FlushToDataSource
                Call oDS_PS_QM230M.RemoveRecord(oDS_PS_QM230M.Size - 1)
                oMat02.LoadFromDataSource
                
                If oMat02.RowCount = 0 Then
                    Call PS_QM230_AddMatrixRow02(0)
                Else
                
                    If Trim(oDS_PS_QM230M.GetValue("U_cls", oMat02.RowCount - 1)) <> "" Then
                        Call PS_QM230_AddMatrixRow02(oMat02.RowCount)
                    End If
                    
                End If
            
            ElseIf oLastItemUID01 = "Mat03" Then
        
                For i = 1 To oMat03.VisualRowCount
                    oMat03.Columns("LineNum").Cells(i).Specific.VALUE = i
                Next i
                
                oMat03.FlushToDataSource
                Call oDS_PS_QM230N.RemoveRecord(oDS_PS_QM230N.Size - 1)
                oMat03.LoadFromDataSource
                
                If oMat03.RowCount = 0 Then
                    Call PS_QM230_AddMatrixRow03(0)
                Else
                
                    If Trim(oDS_PS_QM230N.GetValue("U_cls", oMat03.RowCount - 1)) <> "" Then
                        Call PS_QM230_AddMatrixRow03(oMat03.RowCount)
                    End If
                    
                End If
                
            ElseIf oLastItemUID01 = "Mat04" Then
        
                For i = 1 To oMat04.VisualRowCount
                    oMat04.Columns("LineNum").Cells(i).Specific.VALUE = i
                Next i
                
                oMat04.FlushToDataSource
                Call oDS_PS_QM230O.RemoveRecord(oDS_PS_QM230O.Size - 1)
                oMat04.LoadFromDataSource
                
                If oMat04.RowCount = 0 Then
                    Call PS_QM230_AddMatrixRow04(0)
                Else
                
                    If Trim(oDS_PS_QM230O.GetValue("U_cls", oMat04.RowCount - 1)) <> "" Then
                        Call PS_QM230_AddMatrixRow04(oMat04.RowCount)
                    End If
                    
                End If
            
            ElseIf oLastItemUID01 = "Mat05" Then
        
                For i = 1 To oMat05.VisualRowCount
                    oMat05.Columns("LineNum").Cells(i).Specific.VALUE = i
                Next i
                
                oMat05.FlushToDataSource
                Call oDS_PS_QM230P.RemoveRecord(oDS_PS_QM230P.Size - 1)
                oMat05.LoadFromDataSource
                
                If oMat05.RowCount = 0 Then
                    Call PS_QM230_AddMatrixRow05(0)
                Else
                
                    If Trim(oDS_PS_QM230P.GetValue("U_cls", oMat05.RowCount - 1)) <> "" Then
                        Call PS_QM230_AddMatrixRow05(oMat05.RowCount)
                    End If
                    
                End If
                
            ElseIf oLastItemUID01 = "Mat06" Then
        
                For i = 1 To oMat06.VisualRowCount
                    oMat06.Columns("LineNum").Cells(i).Specific.VALUE = i
                Next i
                
                oMat06.FlushToDataSource
                Call oDS_PS_QM230Q.RemoveRecord(oDS_PS_QM230Q.Size - 1)
                oMat06.LoadFromDataSource
                
                If oMat06.RowCount = 0 Then
                    Call PS_QM230_AddMatrixRow06(0)
                Else
                
                    If Trim(oDS_PS_QM230Q.GetValue("U_cls", oMat06.RowCount - 1)) <> "" Then
                        Call PS_QM230_AddMatrixRow06(oMat06.RowCount)
                    End If
                    
                End If
            
            End If
            
        End If
        
    End If
    
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Function PS_QM230_CreateItems() As Boolean
On Error GoTo PS_QM230_CreateItems_Error
   
    Set oDS_PS_QM230H = oForm01.DataSources.DBDataSources("@PS_QM230H")
    
    Set oDS_PS_QM230L = oForm01.DataSources.DBDataSources("@PS_QM230L")
    Set oDS_PS_QM230M = oForm01.DataSources.DBDataSources("@PS_QM230M")
    Set oDS_PS_QM230N = oForm01.DataSources.DBDataSources("@PS_QM230N")
    Set oDS_PS_QM230O = oForm01.DataSources.DBDataSources("@PS_QM230O")
    Set oDS_PS_QM230P = oForm01.DataSources.DBDataSources("@PS_QM230P")
    Set oDS_PS_QM230Q = oForm01.DataSources.DBDataSources("@PS_QM230Q")
    
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
    
    Exit Function
PS_QM230_CreateItems_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM230_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub PS_QM230_ComboBox_Setting()
On Error GoTo PS_QM230_ComboBox_Setting_Error

'    '사업장 콤보
'    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", False, False)
'
    Dim sQry As String
    sQry = "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'Q230'"

    '구분(Matrix)
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("cls"), sQry)
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat02.Columns("cls"), sQry)
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat03.Columns("cls"), sQry)
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat04.Columns("cls"), sQry)
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat05.Columns("cls"), sQry)
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat06.Columns("cls"), sQry)
    
'    '요청구분(Matrix)
'    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("ReqType"), "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'A002'")

    Exit Sub
PS_QM230_ComboBox_Setting_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM230_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_QM230_CF_ChooseFromList()
On Error GoTo PS_QM230_CF_ChooseFromList_Error
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
PS_QM230_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM230_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_QM230_FormItemEnabled()
On Error GoTo PS_QM230_FormItemEnabled_Error

    Call oForm01.Freeze(True)
    
    If (oForm01.Mode = fm_ADD_MODE) Then
        '//각모드에따른 아이템설정
'
'        oForm01.Items("Remark").Click ct_Regular
        oForm01.Items("DocEntry").Enabled = False
        
        oForm01.Items("StdYM").Enabled = True
        oForm01.Items("ItemCode").Enabled = True
        
        oForm01.Items("Mat01").Enabled = True
        oForm01.Items("Mat02").Enabled = True
        oForm01.Items("Mat03").Enabled = True
        oForm01.Items("Mat04").Enabled = True
        oForm01.Items("Mat05").Enabled = True
        oForm01.Items("Mat06").Enabled = True
        
        Call PS_QM230_FormClear '//UDO방식
        
        Call oForm01.EnableMenu("1281", True) '//찾기
        Call oForm01.EnableMenu("1282", False) '//추가
        
        '년월
        oForm01.Items("StdYM").Specific.VALUE = Format(Now, "YYYY-MM")
        
        '등록일자
        oForm01.Items("DocDate").Specific.VALUE = Format(Now, "YYYYMMDD")
'        '사용자별 사업장 세팅
'        Call oForm01.Items("BPLId").Specific.Select(MDC_PS_Common.User_BPLId(), psk_ByValue)
        '등록자 세팅
'        Call oDS_PS_QM230H.setValue("U_CntcCode", 0, MDC_PS_Common.User_MSTCOD) '데이터소스로 바인딩하면 이벤트 실행 안됨
        oForm01.Items("CntcCode").Specific.VALUE = MDC_PS_Common.User_MSTCOD
        
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//각모드에따른 아이템설정
        
        oForm01.Items("DocEntry").Specific.VALUE = ""
        oForm01.Items("DocEntry").Enabled = True
        
        oForm01.Items("StdYM").Enabled = True
        oForm01.Items("ItemCode").Enabled = True
        
        oForm01.Items("Mat01").Enabled = False
        oForm01.Items("Mat02").Enabled = False
        oForm01.Items("Mat03").Enabled = False
        oForm01.Items("Mat04").Enabled = False
        oForm01.Items("Mat05").Enabled = False
        oForm01.Items("Mat06").Enabled = False
        
        Call oForm01.EnableMenu("1281", False) '//찾기
        Call oForm01.EnableMenu("1282", True) '//추가
        
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        '//각모드에따른 아이템설정
        
        oForm01.Items("DocEntry").Enabled = False
        oForm01.Items("StdYM").Enabled = False
        oForm01.Items("ItemCode").Enabled = False
        
        oForm01.Items("Mat01").Enabled = True
        oForm01.Items("Mat02").Enabled = True
        oForm01.Items("Mat03").Enabled = True
        oForm01.Items("Mat04").Enabled = True
        oForm01.Items("Mat05").Enabled = True
        oForm01.Items("Mat06").Enabled = True
        
    End If
    
    Call oForm01.Freeze(False)
    
    Exit Sub
    
PS_QM230_FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_QM230_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_QM230_AddMatrixRow01(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_QM230_AddMatrixRow01_Error
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//행추가여부
        oDS_PS_QM230L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_QM230L.Offset = oRow
    oDS_PS_QM230L.setValue "U_LineNum", oRow, oRow + 1
    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_QM230_AddMatrixRow01_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_QM230_AddMatrixRow01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_QM230_AddMatrixRow02(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_QM230_AddMatrixRow02_Error
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//행추가여부
        oDS_PS_QM230M.InsertRecord (oRow)
    End If
    oMat02.AddRow
    oDS_PS_QM230M.Offset = oRow
    oDS_PS_QM230M.setValue "U_LineNum", oRow, oRow + 1
    oMat02.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_QM230_AddMatrixRow02_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_QM230_AddMatrixRow02_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_QM230_AddMatrixRow03(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_QM230_AddMatrixRow03_Error
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//행추가여부
        oDS_PS_QM230N.InsertRecord (oRow)
    End If
    oMat03.AddRow
    oDS_PS_QM230N.Offset = oRow
    oDS_PS_QM230N.setValue "U_LineNum", oRow, oRow + 1
    oMat03.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_QM230_AddMatrixRow03_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_QM230_AddMatrixRow03_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_QM230_AddMatrixRow04(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_QM230_AddMatrixRow04_Error
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//행추가여부
        oDS_PS_QM230O.InsertRecord (oRow)
    End If
    oMat04.AddRow
    oDS_PS_QM230O.Offset = oRow
    oDS_PS_QM230O.setValue "U_LineNum", oRow, oRow + 1
    oMat04.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_QM230_AddMatrixRow04_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_QM230_AddMatrixRow04_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_QM230_AddMatrixRow05(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_QM230_AddMatrixRow05_Error
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//행추가여부
        oDS_PS_QM230P.InsertRecord (oRow)
    End If
    oMat05.AddRow
    oDS_PS_QM230P.Offset = oRow
    oDS_PS_QM230P.setValue "U_LineNum", oRow, oRow + 1
    oMat05.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_QM230_AddMatrixRow05_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_QM230_AddMatrixRow05_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_QM230_AddMatrixRow06(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_QM230_AddMatrixRow06_Error
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//행추가여부
        oDS_PS_QM230Q.InsertRecord (oRow)
    End If
    oMat06.AddRow
    oDS_PS_QM230Q.Offset = oRow
    oDS_PS_QM230Q.setValue "U_LineNum", oRow, oRow + 1
    oMat06.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
PS_QM230_AddMatrixRow06_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_QM230_AddMatrixRow06_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_QM230_FormClear()
On Error GoTo PS_QM230_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_QM230'", "")
    If DocEntry = "" Or DocEntry = "0" Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PS_QM230_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM230_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_QM230_EnableMenus()
On Error GoTo PS_QM230_EnableMenus_Error
    '//메뉴활성화
'    Call oForm01.EnableMenu("1288", True)
'    Call oForm01.EnableMenu("1289", True)
'    Call oForm01.EnableMenu("1290", True)
'    Call oForm01.EnableMenu("1291", True)
    '//Call MDC_GP_EnableMenus(oForm01, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False) '//메뉴설정
    Call MDC_GP_EnableMenus(oForm01, False, False, True, True, False, True, True, True, True, False, False, False, False, False, False) '//메뉴설정
    Exit Sub
PS_QM230_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM230_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_QM230_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_QM230_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_QM230_FormItemEnabled
        Call PS_QM230_AddMatrixRow01(0, True) '//UDO방식일때
        Call PS_QM230_AddMatrixRow02(0, True) '//UDO방식일때
        Call PS_QM230_AddMatrixRow03(0, True) '//UDO방식일때
        Call PS_QM230_AddMatrixRow04(0, True) '//UDO방식일때
        Call PS_QM230_AddMatrixRow05(0, True) '//UDO방식일때
        Call PS_QM230_AddMatrixRow06(0, True) '//UDO방식일때
    Else
'        oForm01.Mode = fm_FIND_MODE
'        Call PS_QM230_FormItemEnabled
'        oForm01.Items("DocEntry").Specific.VALUE = oFromDocEntry01
'        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_QM230_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM230_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Function PS_QM230_DataValidCheck() As Boolean
On Error GoTo PS_QM230_DataValidCheck_Error

    PS_QM230_DataValidCheck = False
    Dim i As Long
    
    If (oForm01.Mode = fm_ADD_MODE) Then
        Call PS_QM230_FormClear
    End If
    
    '년월 미입력 시
    If oForm01.Items("StdYM").Specific.VALUE = "" Then
        Call Sbo_Application.SetStatusBarMessage("년월을 입력하지 않았습니다.", bmt_Short, True)
        PS_QM230_DataValidCheck = False
        Exit Function
    End If
    
    '등록자 미입력 시
    If oForm01.Items("CntcCode").Specific.VALUE = "" Then
        Call Sbo_Application.SetStatusBarMessage("등록자를 입력하지 않았습니다.", bmt_Short, True)
        PS_QM230_DataValidCheck = False
        Exit Function
    End If
    
    '품목 미입력 시
    If oForm01.Items("ItemCode").Specific.VALUE = "" Then
        Call Sbo_Application.SetStatusBarMessage("품목을 입력하지 않았습니다.", bmt_Short, True)
        PS_QM230_DataValidCheck = False
        Exit Function
    End If
    
    '등록일자 미입력시
    If oForm01.Items("DocDate").Specific.VALUE = "" Then
        Call Sbo_Application.SetStatusBarMessage("등록일자를 입력하지 않았습니다.", bmt_Short, True)
        PS_QM230_DataValidCheck = False
        Exit Function
    End If

    Call oMat01.FlushToDataSource
    If oMat01.VisualRowCount = 1 Then
        Sbo_Application.SetStatusBarMessage "외경(MAX)의 라인이 존재하지 않습니다.", bmt_Short, True
        PS_QM230_DataValidCheck = False
        Exit Function
    End If
    
    Call oMat02.FlushToDataSource
    If oMat02.VisualRowCount = 1 Then
        Sbo_Application.SetStatusBarMessage "외경(MIN)의 라인이 존재하지 않습니다.", bmt_Short, True
        PS_QM230_DataValidCheck = False
        Exit Function
    End If
    
    Call oMat03.FlushToDataSource
    If oMat03.VisualRowCount = 1 Then
        Sbo_Application.SetStatusBarMessage "내경(MAX)의 라인이 존재하지 않습니다.", bmt_Short, True
        PS_QM230_DataValidCheck = False
        Exit Function
    End If
    
    Call oMat04.FlushToDataSource
    If oMat04.VisualRowCount = 1 Then
        Sbo_Application.SetStatusBarMessage "내경(MIN)의 라인이 존재하지 않습니다.", bmt_Short, True
        PS_QM230_DataValidCheck = False
        Exit Function
    End If
    
    Call oMat05.FlushToDataSource
    If oMat05.VisualRowCount = 1 Then
        Sbo_Application.SetStatusBarMessage "높이(MAX)의 라인이 존재하지 않습니다.", bmt_Short, True
        PS_QM230_DataValidCheck = False
        Exit Function
    End If
    
    Call oMat06.FlushToDataSource
    If oMat06.VisualRowCount = 1 Then
        Sbo_Application.SetStatusBarMessage "높이(MIN)의 라인이 존재하지 않습니다.", bmt_Short, True
        PS_QM230_DataValidCheck = False
        Exit Function
    End If
    
    '외경Max
    Call oDS_PS_QM230L.RemoveRecord(oDS_PS_QM230L.Size - 1)
    Call oMat01.LoadFromDataSource
    
    '외경Min
    Call oDS_PS_QM230M.RemoveRecord(oDS_PS_QM230M.Size - 1)
    Call oMat02.LoadFromDataSource
    
    '내경Max
    Call oDS_PS_QM230N.RemoveRecord(oDS_PS_QM230N.Size - 1)
    Call oMat03.LoadFromDataSource
    
    '내경Min
    Call oDS_PS_QM230O.RemoveRecord(oDS_PS_QM230O.Size - 1)
    Call oMat04.LoadFromDataSource
    
    '높이Max
    Call oDS_PS_QM230P.RemoveRecord(oDS_PS_QM230P.Size - 1)
    Call oMat05.LoadFromDataSource

    '높이Min
    Call oDS_PS_QM230Q.RemoveRecord(oDS_PS_QM230Q.Size - 1)
    Call oMat06.LoadFromDataSource

    If (oForm01.Mode = fm_ADD_MODE) Then
        Call PS_QM230_FormClear
    End If
    
    PS_QM230_DataValidCheck = True
    Exit Function
PS_QM230_DataValidCheck_Error:
    PS_QM230_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PS_QM230_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PS_QM230_MTX01()
On Error GoTo PS_QM230_MTX01_Error
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
        GoTo PS_QM230_MTX01_Exit
    End If
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, False)
    
    For i = 0 To RecordSet01.RecordCount - 1
        If i <> 0 Then
            oDS_PS_QM230L.InsertRecord (i)
        End If
        oDS_PS_QM230L.Offset = i
        Call oDS_PS_QM230L.setValue("U_COL01", i, RecordSet01.Fields(0).VALUE)
        Call oDS_PS_QM230L.setValue("U_COL02", i, RecordSet01.Fields(1).VALUE)
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
PS_QM230_MTX01_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Exit Sub
PS_QM230_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_QM230_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_QM230_DI_API() As Boolean
On Error GoTo PS_QM230_DI_API_Error
    PS_QM230_DI_API = True
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
        ItemInformation(ItemInformationCount).ItemCode = oMat01.Columns("ItemCode").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).BatchNum = oMat01.Columns("BatchNum").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).Quantity = oMat01.Columns("Quantity").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).OPORNo = oMat01.Columns("OPORNo").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).POR1No = oMat01.Columns("POR1No").Cells(i).Specific.VALUE
        ItemInformation(ItemInformationCount).Check = False
        ItemInformationCount = ItemInformationCount + 1
    Next
    
    LineNumCount = 0
    Set oDIObject = Sbo_Company.GetBusinessObject(oPurchaseDeliveryNotes)
    oDIObject.BPL_IDAssignedToInvoice = Trim(oForm01.Items("BPLId").Specific.Selected.VALUE)
    oDIObject.CardCode = Trim(oForm01.Items("CardCode").Specific.VALUE)
    oDIObject.DocDate = Format(oForm01.Items("InDate").Specific.VALUE, "&&&&-&&-&&")
    For i = 0 To ItemInformationCount - 1
        If ItemInformation(i).Check = True Then
            GoTo Continue_First
        End If
        If i <> 0 Then
            oDIObject.Lines.Add
        End If
        oDIObject.Lines.ItemCode = ItemInformation(i).ItemCode
        oDIObject.Lines.WarehouseCode = Trim(oForm01.Items("WhsCode").Specific.VALUE)
        oDIObject.Lines.BaseType = "22"
        oDIObject.Lines.BaseEntry = ItemInformation(i).OPORNo
        oDIObject.Lines.BaseLine = ItemInformation(i).POR1No
        For j = i To UBound(ItemInformation)
            If ItemInformation(j).Check = True Then
                GoTo Continue_Second
            End If
            If (ItemInformation(i).ItemCode <> ItemInformation(j).ItemCode Or ItemInformation(i).OPORNo <> ItemInformation(j).OPORNo Or ItemInformation(i).POR1No <> ItemInformation(j).POR1No) Then
                GoTo Continue_Second
            End If
            '//같은것
            oDIObject.Lines.Quantity = oDIObject.Lines.Quantity + ItemInformation(j).Quantity
            oDIObject.Lines.BatchNumbers.BatchNumber = ItemInformation(j).BatchNum
            oDIObject.Lines.BatchNumbers.Quantity = ItemInformation(j).Quantity
            oDIObject.Lines.BatchNumbers.Add
            ItemInformation(j).PDN1No = LineNumCount
            ItemInformation(j).Check = True
Continue_Second:
        Next
        LineNumCount = LineNumCount + 1
Continue_First:
    Next
    RetVal = oDIObject.Add
    If RetVal = 0 Then
        ResultDocNum = Sbo_Company.GetNewObjectKey
        For i = 0 To UBound(ItemInformation)
            Call oDS_PS_QM230L.setValue("U_OPDNNo", i, ResultDocNum)
            Call oDS_PS_QM230L.setValue("U_PDN1No", i, ItemInformation(i).PDN1No)
        Next
    Else
        GoTo PS_QM230_DI_API_Error
    End If
    
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_Commit
    End If
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    
    Set oDIObject = Nothing
    Exit Function
PS_QM230_DI_API_DI_Error:
    If Sbo_Company.InTransaction = True Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Application.SetStatusBarMessage Sbo_Company.GetLastErrorCode & " - " & Sbo_Company.GetLastErrorDescription, bmt_Short, True
    PS_QM230_DI_API = False
    Set oDIObject = Nothing
    Exit Function
PS_QM230_DI_API_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM230_DI_API_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    PS_QM230_DI_API = False
End Function

Private Sub PS_QM230_FormResize()
On Error GoTo PS_QM230_FormResize_Error
    
    '그룹박스 크기 동적 할당
    oForm01.Items("Rec01").Height = oForm01.Items("Mat01").Height + 53
    oForm01.Items("Rec01").Width = oForm01.Items("Mat01").Width + 20
    
    Exit Sub
PS_QM230_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM230_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PS_QM230_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PS_QM230_Validate_Error
    PS_QM230_Validate = True
    Dim i, j As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
'    If MDC_PS_Common.GetValue("SELECT Canceled FROM [@PS_QM230H] WHERE DocEntry = '" & oForm01.Items("DocEntry").Specific.Value & "'", 0, 1) = "Y" Then
'        Call MDC_Com.MDC_GF_Message("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", "W")
'        PS_QM230_Validate = False
'        GoTo PS_QM230_Validate_Exit
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
'                    PS_QM230_Validate = False
'                    GoTo PS_QM230_Validate_Exit
'                End If
'            End If
'            RecordSet01.MoveNext
'        Next
    ElseIf ValidateType = "행삭제" Then
        '//행삭제전 행삭제가능여부검사
'        If oForm01.Mode = fm_OK_MODE Or oForm01.Mode = fm_UPDATE_MODE Then '//추가,수정모드일때행삭제가능검사
'            If (oMat01.Columns("U_LineNum").Cells(oLastColRow01).Specific.Value = "") Then '//새로추가된 행인경우, 삭제하여도 무방하다
'            Else
'                If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_SD030L] WHERE U_ORDRNum = '" & Val(oForm01.Items("8").Specific.Value) & "' AND U_RDR1Num = '" & Val(oMat01.Columns("U_LineNum").Cells(oLastColRow01).Specific.Value) & "'", 0, 1)) > 0 Then
'                    MDC_Com.MDC_GF_Message "이미출하,선출요청된 행입니다. 삭제할수 없습니다.", "W"
'                    PS_QM230_Validate = False
'                    GoTo PS_QM230_Validate_Exit
'                End If
'            End If
'        End If
    ElseIf ValidateType = "취소" Then
'        Query01 = "SELECT DocEntry,LineNum,ItemCode FROM [RDR1] WHERE DocEntry = '" & oForm01.Items("8").Specific.Value & "'"
'        RecordSet01.DoQuery Query01
'        For i = 0 To RecordSet01.RecordCount - 1
'            If (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_SD030L] WHERE U_ORDRNum = '" & Val(RecordSet01.Fields(0).Value) & "' AND U_RDR1Num = '" & Val(RecordSet01.Fields(1).Value) & "'", 0, 1)) > 0 Then
'                MDC_Com.MDC_GF_Message "출하,선출요청된문서입니다. 적용할수 없습니다.", "W"
'                PS_QM230_Validate = False
'                GoTo PS_QM230_Validate_Exit
'            End If
'            RecordSet01.MoveNext
'        Next
    End If
    Set RecordSet01 = Nothing
    Exit Function
PS_QM230_Validate_Exit:
    Set RecordSet01 = Nothing
    Exit Function
PS_QM230_Validate_Error:
    PS_QM230_Validate = False
    Sbo_Application.SetStatusBarMessage "PS_QM230_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Public Sub PS_QM230_OpenGraph(ByVal prmStdYM As String, ByVal prmItemCode As String, ByVal prmDataType As String)
'******************************************************************************
'Function ID : PS_QM230_OpenGraph()
'해당모듈 : PS_QM230
'기능 : VB확장 프로그램을 이용하여 그래프를 출력
'인수 : prmStdYM(년월), prmItemCode(품목코드), prmDataType(측정 부위(PanelLevel), 1:외경Max, 2:외경Min, 3:내경Max, 4:내경Min, 5:높이Max, 6:높이Min)
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_QM230_OpenGraph_Error
    
    Dim Shell As Object '외부 확장 프로그램 실행용 오브젝트
    
    Set Shell = CreateObject("Wscript.Shell")
    Call Shell.Run("C:\BOne_Expansion\PS_QM230_EndBearingGraph.exe " & prmStdYM & "_" & prmItemCode & "_" & prmDataType, 1, True)
'    Call Shell.Run("\\191.1.1.220\B1_SHR\BOne_Expansion\PS_QM230_EndBearingGraph.exe " & prmStdYM & "_" & prmItemCode & "_" & prmDataType, 1, True)
    
    Exit Sub

PS_QM230_OpenGraph_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM230_OpenGraph_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Function PS_QM230_CheckDocument() As Boolean
'******************************************************************************
'Function ID : PS_QM230_CheckDocument()
'해당모듈 : PS_QM230
'기능 : 동일 년월, 동일 검사자, 동일 품목의 데이터가 이미 등록되었는지 검사
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_QM230_CheckDocument_Error

    Dim i As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Query01 = "                 SELECT       COUNT(*)"
    Query01 = Query01 & "  FROM        [@PS_QM230H] AS T0"
    Query01 = Query01 & "  WHERE       T0.[Status] = 'O'"
    Query01 = Query01 & "                  AND T0.U_StdYM = '" & oForm01.Items("StdYM").Specific.VALUE & "'"
    Query01 = Query01 & "                  AND T0.U_CntcCode = '" & oForm01.Items("CntcCode").Specific.VALUE & "'"
    Query01 = Query01 & "                  AND T0.U_ItemCode = '" & oForm01.Items("ItemCode").Specific.VALUE & "'"
    
    Call RecordSet01.DoQuery(Query01)
    
    If RecordSet01.Fields(0).VALUE = 0 Then
        PS_QM230_CheckDocument = True
    Else
        PS_QM230_CheckDocument = False
        Call MDC_Com.MDC_GF_Message("해당 년월에 이미 자료가 등록되었습니다. 확인하십시오.", "E")
    End If
    
    Set RecordSet01 = Nothing

    Exit Function
    
PS_QM230_CheckDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM230_CheckDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function
