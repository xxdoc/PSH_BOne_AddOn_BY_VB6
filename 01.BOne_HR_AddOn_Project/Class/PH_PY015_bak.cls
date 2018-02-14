VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY015"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY015.cls
'//  Module         : 근태관리 > 근태등록
'//  Desc           : 연차이월등록
'//********************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form

Public oGrid1           As SAPbouiCOM.Grid
Public oDS_PH_PY015A     As SAPbouiCOM.DataTable


Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long

Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY015.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY015_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY015"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
'    oForm.DataBrowser.BrowseBy = "Code"
        
    oForm.PaneLevel = 1
    oForm.Freeze True
    Call PH_PY015_CreateItems
    Call PH_PY015_FormItemEnabled
    Call PH_PY015_EnableMenus
'    Call PH_PY015_SetDocument(oFromDocEntry01)
'    Call PH_PY015_FormResize
    
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

Private Function PH_PY015_CreateItems() As Boolean

    Dim sQry        As String
    Dim i           As Long
    Dim CLTCOD      As String
    
    Dim oCheck      As SAPbouiCOM.CheckBox
    Dim oEdit       As SAPbouiCOM.EditText
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns    As SAPbouiCOM.Columns
    Dim optBtn      As SAPbouiCOM.OptionBtn
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY015_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Set oGrid1 = oForm.Items("Grid01").Specific
    
    oGrid1.SelectionMode = ms_Single
    
    oForm.DataSources.DataTables.Add ("PH_PY015")
    
    oGrid1.DataTable = oForm.DataSources.DataTables.Item("PH_PY015")
    Set oDS_PH_PY015A = oForm.DataSources.DataTables.Item("PH_PY015")
    

    '//----------------------------------------------------------------------------------------------
    '// 기본사항
    '//----------------------------------------------------------------------------------------------
    
    '//사업장
    
    Set oCombo = oForm.Items("CLTCOD").Specific
'    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
'    Call SetReDataCombo(oForm, sQry, oCombo)
'
'    CLTCOD = MDC_SetMod.Get_ReData("Branch", "USER_CODE", "OUSR", "'" & oCompany.UserName & "'")
'    oCombo.Select CLTCOD, psk_ByValue
    
       
'    oCombo.Select 0, psk_Index
    oForm.Items("CLTCOD").DisplayDesc = True
    
    '//년도
    Call oForm.DataSources.UserDataSources.Add("Year", dt_SHORT_TEXT, 4)
    oForm.Items("Year").Specific.DataBind.SetBound True, "", "Year"
    
    '//사번
    Call oForm.DataSources.UserDataSources.Add("MSTCOD", dt_SHORT_TEXT, 10)
    oForm.Items("MSTCOD").Specific.DataBind.SetBound True, "", "MSTCOD"
    '//성명
    Call oForm.DataSources.UserDataSources.Add("FullName", dt_SHORT_TEXT, 10)
    oForm.Items("FullName").Specific.DataBind.SetBound True, "", "FullName"
    
    
    '//입사일
    Call oForm.DataSources.UserDataSources.Add("ipsail", dt_SHORT_TEXT, 8)
    oForm.Items("ipsail").Specific.DataBind.SetBound True, "", "ipsail"
    
    '//이월
    Call oForm.DataSources.UserDataSources.Add("iwol", dt_QUANTITY, 10)
    oForm.Items("iwol").Specific.DataBind.SetBound True, "", "iwol"
    
    '//발생연차
    Call oForm.DataSources.UserDataSources.Add("yycha", dt_QUANTITY, 10)
    oForm.Items("yycha").Specific.DataBind.SetBound True, "", "yycha"
    
    '//발생월차
    Call oForm.DataSources.UserDataSources.Add("mmcha", dt_QUANTITY, 10)
    oForm.Items("mmcha").Specific.DataBind.SetBound True, "", "mmcha"
    
    '//발생근속차
    Call oForm.DataSources.UserDataSources.Add("gunsok", dt_QUANTITY, 10)
    oForm.Items("gunsok").Specific.DataBind.SetBound True, "", "gunsok"
    
    '//사용연차
    Call oForm.DataSources.UserDataSources.Add("useyy", dt_QUANTITY, 10)
    oForm.Items("useyy").Specific.DataBind.SetBound True, "", "useyy"
    
    '//사용반차
    Call oForm.DataSources.UserDataSources.Add("usebyy", dt_QUANTITY, 10)
    oForm.Items("usebyy").Specific.DataBind.SetBound True, "", "usebyy"
    
    '//잔여수
    Call oForm.DataSources.UserDataSources.Add("jandd", dt_QUANTITY, 10)
    oForm.Items("jandd").Specific.DataBind.SetBound True, "", "jandd"
    
    '//적치수
    Call oForm.DataSources.UserDataSources.Add("Savedd", dt_QUANTITY, 10)
    oForm.Items("Savedd").Specific.DataBind.SetBound True, "", "Savedd"
    
    '//지급수
    Call oForm.DataSources.UserDataSources.Add("Paydd", dt_QUANTITY, 10)
    oForm.Items("Paydd").Specific.DataBind.SetBound True, "", "Paydd"
        
    
    oForm.Update
    
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
    
PH_PY015_CreateItems_Error:
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY015_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY015_EnableMenus()
    
    On Error GoTo PH_PY015_EnableMenus_Error

    Call oForm.EnableMenu("1283", False)     '//제거
    Call oForm.EnableMenu("1284", False)    '//취소
    Call oForm.EnableMenu("1293", False)     '//행삭제
    
    Exit Sub
    
PH_PY015_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY015_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY015_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY015_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY015_FormItemEnabled
'        Call PH_PY015_AddMatrixRow
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY015_FormItemEnabled
        oForm.Items("Code").Specific.VALUE = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY015_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY015_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY015_FormItemEnabled()
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim sQry        As String
    Dim i           As Long
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    
    Dim CLTCOD As String
    Dim sPosDate As String
    
    On Error GoTo PH_PY015_FormItemEnabled_Error
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Call oForm.Freeze(True)
    If (oForm.Mode = fm_ADD_MODE) Then
    
        
        Call oForm.EnableMenu("1281", False)     '//문서찾기
        Call oForm.EnableMenu("1282", True)    '//문서추가
        
        oForm.Items("Year").Specific.VALUE = Format(Now, "YYYY")
        oForm.Items("MSTCOD").Specific.VALUE = ""
        oForm.Items("FullName").Specific.VALUE = ""
        
            
        oForm.Items("iwol").Specific.VALUE = 0
        oForm.Items("yycha").Specific.VALUE = 0
        oForm.Items("mmcha").Specific.VALUE = 0
        oForm.Items("gunsok").Specific.VALUE = 0
        oForm.Items("useyy").Specific.VALUE = 0
        oForm.Items("usebyy").Specific.VALUE = 0
        oForm.Items("jandd").Specific.VALUE = 0
        
        oForm.Items("Savedd").Specific.VALUE = 0
        oForm.Items("Paydd").Specific.VALUE = 0
        
        oForm.Items("TeamName").Specific.VALUE = ""
        oForm.Items("RspName").Specific.VALUE = ""
        
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")


    ElseIf (oForm.Mode = fm_FIND_MODE) Then
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        
        Call oForm.EnableMenu("1281", False)    '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가
        
        
    ElseIf (oForm.Mode = fm_OK_MODE) Then
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD", False)
        
        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가
        
    End If
    Set oCombo = Nothing
    Call oForm.Freeze(False)
    Exit Sub
    
PH_PY015_FormItemEnabled_Error:
    Set oCombo = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY015_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Public Sub Raise_FormItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
    Dim sQry        As String
    Dim i           As Long
    Dim tSex        As String
    Dim tBrith      As String
    Dim Day         As String
    Dim ActCode     As String
    Dim MSTCOD      As String
    
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns     As SAPbouiCOM.Columns
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo Raise_FormItemEvent_Error
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Select Case pval.EventType
        Case et_ITEM_PRESSED: '//1
            
            If pval.BeforeAction = True Then
                If pval.ItemUID = "1" Then
                    If PH_PY015_DataValidCheck = False Then
                        BubbleEvent = False
                    End If
                End If
                
                If pval.ItemUID = "Btn_ret" Then
                    Call PH_PY015_MTX01
                End If
                
                
                
                If pval.ItemUID = "Btn01" Then
                    Call PH_PY015_SAVE
                End If
                
                                
                If pval.ItemUID = "Btn_del" Then
                    Call PH_PY015_Delete
                End If
'                If oForm.Mode = fm_FIND_MODE Then
'                    If pval.ItemUID = "Btn01" Then
'                        Sbo_Application.ActivateMenuItem ("7425")
'                        BubbleEvent = False
'                    End If
'
'                End If
            ElseIf pval.BeforeAction = False Then
                Select Case pval.ItemUID
                Case "1"
                    If oForm.Mode = fm_ADD_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY015_FormItemEnabled
                        End If
                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY015_FormItemEnabled
                        End If
                    ElseIf oForm.Mode = fm_OK_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY015_FormItemEnabled
                        End If
                    End If
'
                End Select
            End If
     
'----------------------------------------------------------
        Case et_KEY_DOWN: '//2
            If pval.BeforeAction = True Then
                If pval.CharPressed = 9 Then
                    If pval.ItemUID = "SShiftDat" Then
                    End If
                End If
            End If
'----------------------------------------------------------
        Case et_GOT_FOCUS: '//3
            Select Case pval.ItemUID
            Case "Mat01"
                If pval.Row > 0 Then
                    oLastItemUID = pval.ItemUID
                    oLastColUID = pval.ColUID
                    oLastColRow = pval.Row
                End If
            Case Else
                oLastItemUID = pval.ItemUID
                oLastColUID = ""
                oLastColRow = 0
            End Select

'----------------------------------------------------------
        Case et_LOST_FOCUS: '//4

'----------------------------------------------------------
        Case et_COMBO_SELECT: '//5
            Call oForm.Freeze(True)
            If pval.BeforeAction = True Then
                
            ElseIf pval.BeforeAction = False Then
                If pval.ItemChanged = True Then
                    '//사업장(헤더)
                    If pval.ItemUID = "SCLTCOD" Then
                        
                    End If
                    
                End If
            End If

            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_CLICK: '//6
            Call oForm.Freeze(True)
            If pval.BeforeAction = True Then
                Select Case pval.ItemUID
                Case "Grid01"
                    If pval.Row >= 0 Then
                        Select Case pval.ItemUID
                        Case "Grid01"
                            'Call oMat1.SelectRow(pval.Row, True, False)
                            PH_PY015_MTX02 pval.ItemUID, pval.Row, pval.ColUID
                        End Select
                        
                    End If
                End Select
                
                Select Case pval.ItemUID
                Case "Grid01"
                    If pval.Row > 0 Then
                        oLastItemUID = pval.ItemUID
                        oLastColUID = pval.ColUID
                        oLastColRow = pval.Row
                    End If
                Case Else
                    oLastItemUID = pval.ItemUID
                    oLastColUID = ""
                    oLastColRow = 0
                End Select
            ElseIf pval.BeforeAction = False Then

            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_DOUBLE_CLICK: '//7
            Call oForm.Freeze(True)
            If pval.BeforeAction = True Then
            Else
                
            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_MATRIX_LINK_PRESSED '//8

'----------------------------------------------------------
        Case et_MATRIX_COLLAPSE_PRESSED '//9
'----------------------------------------------------------
        Case et_VALIDATE: '//10
'            Call oForm.Freeze(True)
            If pval.BeforeAction = True Then
                If pval.ItemChanged = True Then
                    
                End If
                    
            ElseIf pval.BeforeAction = False Then
                If pval.ItemChanged = True Then
                    Select Case pval.ItemUID

                        Case "Savedd"
                            
                            If Left(oForm.Items("ipsail").Specific.VALUE, 4) = oForm.Items("Year").Specific.VALUE And Right(oForm.Items("ipsail").Specific.VALUE, 4) <> "0101" And oForm.Items("jandd").Specific.VALUE > 0 Then
                                oForm.Items("Paydd").Specific.VALUE = 0
                            Else
                                oForm.Items("Paydd").Specific.VALUE = oForm.Items("jandd").Specific.VALUE - oForm.Items("Savedd").Specific.VALUE
                            
                            End If
                        
                    End Select
                 
                End If
            End If
'            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_MATRIX_LOAD: '//11
            If pval.BeforeAction = True Then
            ElseIf pval.BeforeAction = False Then
'                oMat1.LoadFromDataSource
'                Call PH_PY015_AddMatrixRow
                
            End If
'----------------------------------------------------------
        Case et_DATASOURCE_LOAD '//12
           
'----------------------------------------------------------
        Case et_FORM_LOAD: '//16
            
'----------------------------------------------------------
        Case et_FORM_UNLOAD: '//17
            If pval.BeforeAction = True Then
            ElseIf pval.BeforeAction = False Then
                RemoveForms oFormUniqueID
                Set oForm = Nothing
                Set oDS_PH_PY015A = Nothing
                
'                Set oMat1 = Nothing
            End If
'----------------------------------------------------------
        Case et_FORM_ACTIVATE: '//18
            
'----------------------------------------------------------
        Case et_FORM_DEACTIVATE: '//19
            
'----------------------------------------------------------
        Case et_FORM_CLOSE '//20
            
'----------------------------------------------------------
        Case et_FORM_RESIZE '//21
            If pval.BeforeAction = True Then
        
            ElseIf pval.BeforeAction = False Then
                
            End If
'----------------------------------------------------------
        Case et_FORM_KEY_DOWN '//22
            
'----------------------------------------------------------
        Case et_FORM_MENU_HILIGHT '//23
            
'----------------------------------------------------------
        Case et_CHOOSE_FROM_LIST '//27
            If pval.BeforeAction = True Then
            
            ElseIf pval.Before_Action = False Then
'                If pval.ItemUID = "Code" Then
'                    Call MDC_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY015A", "Code")
'                End If
            End If

'----------------------------------------------------------
        Case et_PICKER_CLICKED '//37
            
'----------------------------------------------------------
        Case et_GRID_SORT '//38
            
'----------------------------------------------------------
        Case et_Drag '//39
        
            
    End Select
    
    Set oCombo = Nothing
    Set oRecordSet = Nothing
    
    Exit Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Raise_FormItemEvent_Error:
    oForm.Freeze (False)
    Set oCombo = Nothing
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "Raise_FormItemEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Public Sub Raise_FormMenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
    Dim i As Long
    On Error GoTo Raise_FormMenuEvent_Error
    Call oForm.Freeze(True)
    
    If (pval.BeforeAction = True) Then
        Select Case pval.MenuUID
            Case "1283":
                If Sbo_Application.MessageBox("현재 화면내용전체를 제거 하시겠습니까? 복구할 수 없습니다.", 2, "Yes", "No") = 2 Then
                    BubbleEvent = False
                    Exit Sub
                End If
            Case "1284":
            Case "1286":
            Case "1293":
            Case "1281":
            Case "1282":
            Case "1288", "1289", "1290", "1291":
'                Call PH_PY015_FormItemEnabled
        End Select
    ElseIf (pval.BeforeAction = False) Then
        Select Case pval.MenuUID
            Case "1283":
                oForm.Mode = fm_ADD_MODE
                Call PH_PY015_FormItemEnabled
'                Call PH_PY015_AddMatrixRow
            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//문서찾기
                Call PH_PY015_FormItemEnabled
'                Call PH_PY015_AddMatrixRow
                oForm.Items("Code").CLICK ct_Regular
            Case "1282":    '//문서추가
                Call PH_PY015_FormItemEnabled
'                Call PH_PY015_AddMatrixRow
            Case "1288", "1289", "1290", "1291":
                Call PH_PY015_FormItemEnabled
            Case "1293" '// 행삭제
                
        End Select
    End If
    Call oForm.Freeze(False)
    Exit Sub
Raise_FormMenuEvent_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_MenuEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub Raise_FormDataEvent(ByRef FormUID As String, ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
    Dim i       As Long
    Dim sQry    As String
    Dim oCombo  As SAPbouiCOM.ComboBox
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    
    On Error GoTo Raise_FormDataEvent_Error
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If (BusinessObjectInfo.BeforeAction = False) Then
        Select Case BusinessObjectInfo.EventType
            Case et_FORM_DATA_LOAD:     '//33
            Case et_FORM_DATA_ADD:      '//34
            Case et_FORM_DATA_UPDATE:   '//35
            Case et_FORM_DATA_DELETE:   '//36
        End Select

    End If
    Set oCombo = Nothing
    Set oRecordSet = Nothing
    Exit Sub
    
Raise_FormDataEvent_Error:
    Set oCombo = Nothing
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "Raise_FormDataEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True

End Sub

Public Sub Raise_RightClickEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
    
    On Error GoTo Raise_RightClickEvent_Error
    
    If pval.BeforeAction = True Then
    ElseIf pval.BeforeAction = False Then
    End If
    Select Case pval.ItemUID
    Case "Mat01"
        If pval.Row > 0 Then
            oLastItemUID = pval.ItemUID
            oLastColUID = pval.ColUID
            oLastColRow = pval.Row
        End If
    Case Else
        oLastItemUID = pval.ItemUID
        oLastColUID = ""
        oLastColRow = 0
    End Select
    Exit Sub
Raise_RightClickEvent_Error:

    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Sub PH_PY015_FormClear()
On Error GoTo PH_PY015_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY015'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PH_PY015_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY015_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY015_DataValidCheck() As Boolean
On Error GoTo PH_PY015_DataValidCheck_Error
    PH_PY015_DataValidCheck = False
    Dim i       As Long
    Dim j       As Long
    
    Dim sQry    As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Exit Function
    
    Set oRecordSet = Nothing
    
    
PH_PY015_DataValidCheck_Error:
    Set oRecordSet = Nothing
    PH_PY015_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY015_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY015_MTX01()

    '//메트릭스에 데이터 로드
    
    Dim i       As Long
    Dim sQry    As String
    Dim iRow        As Long
    
    Dim Param01 As String
    Dim Param02 As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY015_MTX01_Error

    Call oForm.Freeze(True)
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    Param01 = oForm.Items("CLTCOD").Specific.VALUE
    Param02 = oForm.Items("Year").Specific.VALUE
    
    If Trim(Param01) = "" Then
        Call MDC_Com.MDC_GF_Message("사업장이 없습니다. 확인바랍니다..", "E")
        GoTo PH_PY015_MTX01_Exit
    End If
        
    If Trim(Param02) = "" Then
        Call MDC_Com.MDC_GF_Message("년도가 없습니다. 확인바랍니다..", "E")
        GoTo PH_PY015_MTX01_Exit
    End If
        
        
        
    sQry = "EXEC PH_PY015_01 '" & Param01 & "', '" & Param02 & "'"

    Call oDS_PH_PY015A.ExecuteQuery(sQry)
    
    
    
    iRow = oForm.DataSources.DataTables.Item(0).Rows.Count
    
    Call PH_PY015_TitleSetting(iRow)
    
    oForm.Update
    
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Sub
PH_PY015_MTX01_Exit:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Sub
PH_PY015_MTX01_Error:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY015_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub
Private Sub PH_PY015_MTX02(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)


    '//그리드 자료를 head에 로드
    
    Dim i       As Long
    Dim sQry    As String
    Dim sRow As Long
    
    Dim Param01 As String
    Dim Param02 As String
    Dim Param03 As String
    
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY015_MTX02_Error

    Call oForm.Freeze(True)
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    sRow = oRow
    
    
    Param01 = oForm.Items("CLTCOD").Specific.VALUE
    Param02 = oDS_PH_PY015A.Columns.Item("Year").Cells(oRow).VALUE
    Param03 = oDS_PH_PY015A.Columns.Item("MSTCOD").Cells(oRow).VALUE
    
    
    
    sQry = "EXEC PH_PY775_01 '" & Param01 & "', '" & Param02 & "', '" & Param03 & "'"
    Call oRecordSet.DoQuery(sQry)
    
    If (oRecordSet.RecordCount = 0) Then
        
        oForm.Items("MSTCOD").Specific.VALUE = oDS_PH_PY015A.Columns.Item("MSTCOD").Cells(oRow).VALUE
        oForm.Items("FullName").Specific.VALUE = oDS_PH_PY015A.Columns.Item("FullName").Cells(oRow).VALUE
        
            
        oForm.Items("iwol").Specific.VALUE = 0
        oForm.Items("yycha").Specific.VALUE = 0
        oForm.Items("mmcha").Specific.VALUE = 0
        oForm.Items("gunsok").Specific.VALUE = 0
        oForm.Items("useyy").Specific.VALUE = 0
        oForm.Items("usebyy").Specific.VALUE = 0
        oForm.Items("jandd").Specific.VALUE = 0
        
        oForm.Items("Savedd").Specific.VALUE = 0
        oForm.Items("Paydd").Specific.VALUE = 0
        
        oForm.Items("TeamName").Specific.VALUE = oRecordSet.Fields("TeamName").VALUE
        oForm.Items("RspName").Specific.VALUE = oRecordSet.Fields("RspName").VALUE
    
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
        GoTo PH_PY015_MTX02_Exit
    End If
    
    
'    oForm.Items("PosDate").Specific.Value = Format(oRecordset.Fields("PosDate").Value, "YYYYMMDD")
        
'    oForm.DataSources.UserDataSources.Item("PosDate").Value = Format(oRecordSet.Fields("PosDate").Value, "YYYYMMDD")
        
        
    oForm.Items("MSTCOD").Specific.VALUE = oRecordSet.Fields("MSTCOD").VALUE
    oForm.Items("FullName").Specific.VALUE = oRecordSet.Fields("FullName").VALUE
    oForm.Items("ipsail").Specific.VALUE = Format(oRecordSet.Fields("ipsail").VALUE, "YYYYMMDD")
        
    oForm.Items("iwol").Specific.VALUE = oRecordSet.Fields("iwol").VALUE
    oForm.Items("yycha").Specific.VALUE = oRecordSet.Fields("yycha").VALUE
    oForm.Items("mmcha").Specific.VALUE = oRecordSet.Fields("mmcha").VALUE
    oForm.Items("gunsok").Specific.VALUE = oRecordSet.Fields("gunsok").VALUE
    oForm.Items("useyy").Specific.VALUE = oRecordSet.Fields("useyy").VALUE
    oForm.Items("usebyy").Specific.VALUE = oRecordSet.Fields("usebyy").VALUE
    oForm.Items("jandd").Specific.VALUE = oRecordSet.Fields("jandd").VALUE
        
    If Left(oForm.Items("ipsail").Specific.VALUE, 4) = oForm.Items("Year").Specific.VALUE And Right(oForm.Items("ipsail").Specific.VALUE, 4) <> "0101" Then
        oForm.Items("Savedd").Specific.VALUE = 0
        oForm.Items("Paydd").Specific.VALUE = 0
    Else
        oForm.Items("Savedd").Specific.VALUE = oRecordSet.Fields("Savedd").VALUE
        oForm.Items("Paydd").Specific.VALUE = oRecordSet.Fields("Paydd").VALUE
    End If
    
    
    oForm.Items("TeamName").Specific.VALUE = oRecordSet.Fields("TeamName").VALUE
    oForm.Items("RspName").Specific.VALUE = oRecordSet.Fields("RspName").VALUE
    
    
    '//신규일때 잔여일수를 적치일수에 표시
    If oRecordSet.Fields("jandd").VALUE > 0 And oRecordSet.Fields("Savedd").VALUE = 0 And oRecordSet.Fields("Paydd").VALUE = 0 Then
        If Left(oForm.Items("ipsail").Specific.VALUE, 4) = oForm.Items("Year").Specific.VALUE And Right(oForm.Items("ipsail").Specific.VALUE, 2) <> "0101" Then
            oForm.Items("Savedd").Specific.VALUE = 0
        Else
            oForm.Items("Savedd").Specific.VALUE = oRecordSet.Fields("jandd").VALUE
        End If
    End If
    
    
    
'    '//부서
    
    
    oForm.Update
    
    
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Sub
PH_PY015_MTX02_Exit:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    
    Exit Sub
PH_PY015_MTX02_Error:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY015_MTX02_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY015_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY015_Validate_Error
    PH_PY015_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY015A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", bmt_Short, True
        PH_PY015_Validate = False
        GoTo PH_PY015_Validate_Exit
    End If
'
    If ValidateType = "수정" Then

    ElseIf ValidateType = "행삭제" Then

    ElseIf ValidateType = "취소" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY015_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY015_Validate_Error:
    PH_PY015_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY015_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

'//행삭제 (FormUID, pval, BubbleEvent, 매트릭스 이름, 디비데이터소스, 데이터 체크 필드명)
Private Sub Raise_EVENT_ROW_DELETE(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean, _
                                                    oMat As SAPbouiCOM.Matrix, DBData As SAPbouiCOM.DBDataSource, CheckField As String)

    Dim i As Long
    
    On Error GoTo Raise_EVENT_ROW_DELETE_Error

    If (oLastColRow > 0) Then
        If pval.BeforeAction = True Then
        
        ElseIf pval.BeforeAction = False Then
            If oMat.RowCount <> oMat.VisualRowCount Then
                oMat.FlushToDataSource
    
                While (i <= DBData.Size - 1)
                    If DBData.GetValue(CheckField, i) = "" Then
                        DBData.RemoveRecord (i)
                        i = 0
                    Else
                        i = i + 1
                    End If
                Wend
                
                For i = 0 To DBData.Size
                    Call DBData.setValue("U_LineNum", i, i + 1)
                Next i
                
                oMat.LoadFromDataSource
            End If
        End If
    End If
Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub PH_PY015_SAVE()

    '//데이타 저장
    
    Dim i       As Long
    Dim sQry    As String
    
    Dim CLTCOD$, MSTCOD$, FullName$, YEAR$
    
    Dim Savedd As Double
    Dim Paydd As Double
    Dim Usedd As Double
    Dim Usebdd As Double
    
    Dim DocEntry As Integer
    
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY015_SAVE_Error

    Call oForm.Freeze(True)
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    CLTCOD = oForm.Items("CLTCOD").Specific.VALUE
    MSTCOD = oForm.Items("MSTCOD").Specific.VALUE
    FullName = oForm.Items("FullName").Specific.VALUE

    YEAR = oForm.Items("Year").Specific.VALUE
    
    If Trim(YEAR) = "" Then
        Call MDC_Com.MDC_GF_Message("년도가 없습니다. 확인바랍니다..", "E")
        GoTo PH_PY015_SAVE_Exit
    End If
    
    If Trim(CLTCOD) = "" Then
        Call MDC_Com.MDC_GF_Message("사업장이 없습니다. 확인바랍니다..", "E")
        GoTo PH_PY015_SAVE_Exit
    End If
    If Trim(MSTCOD) = "" Then
        Call MDC_Com.MDC_GF_Message("사번이 없습니다. 확인바랍니다..", "E")
        GoTo PH_PY015_SAVE_Exit
    End If
    
    
    
    
    Savedd = oForm.Items("Savedd").Specific.VALUE
    Paydd = oForm.Items("Paydd").Specific.VALUE
    Usedd = oForm.Items("useyy").Specific.VALUE
    Usebdd = oForm.Items("usebyy").Specific.VALUE
        
    sQry = " Select Count(*) From [@PH_PY015A] Where U_CLTCOD = '" & CLTCOD & "' And U_Year = '" & YEAR & "' And U_MSTCOD = '" & MSTCOD & "'"
    oRecordSet.DoQuery sQry
    
    If oRecordSet.Fields(0).VALUE > 0 Then
        '//갱신
        
        sQry = "Update [@PH_PY015A] set U_Savedd = " & Savedd & ", U_Paydd = " & Paydd & ",  U_Usedd = " & Usedd
        sQry = sQry + " Where U_CLTCOD = '" & CLTCOD & "' And U_Year = '" & YEAR & "' And U_MSTCOD = '" & MSTCOD & "'"
        
        oRecordSet.DoQuery sQry
        
    Else
        sQry = "Select isnull(Max(DocEntry),0) + 1 From [@PH_PY015A]"
        oRecordSet.DoQuery sQry
        DocEntry = oRecordSet.Fields(0).VALUE
        
        '//신규
        sQry = "INSERT INTO [@PH_PY015A]"
                            sQry = sQry & " ("
                            sQry = sQry & " DocEntry,"
                            sQry = sQry & " DocNum,"
                            sQry = sQry & " U_CLTCOD,"
                            sQry = sQry & " U_Year,"
                            sQry = sQry & " U_MSTCOD,"
                            sQry = sQry & " U_FullName,"
                            sQry = sQry & " U_Savedd,"
                            sQry = sQry & " U_Paydd,"
                            sQry = sQry & " U_Usedd"
                            sQry = sQry & " ) "
                            sQry = sQry & "VALUES("
                            sQry = sQry & DocEntry & ","
                            sQry = sQry & DocEntry & ","
                            sQry = sQry & "'" & CLTCOD & "',"
                            sQry = sQry & "'" & YEAR & "',"
                            sQry = sQry & "'" & MSTCOD & "',"
                            sQry = sQry & "'" & FullName & "',"
                            sQry = sQry & Savedd & ","
                            sQry = sQry & Paydd & ","
                            sQry = sQry & Usedd + Usebdd & " )"
                            oRecordSet.DoQuery sQry
    End If
    
    



    Set oRecordSet = Nothing
    Call oForm.Freeze(False)

    Call PH_PY015_MTX01

    Exit Sub
    
PH_PY015_SAVE_Exit:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    
    Exit Sub
PH_PY015_SAVE_Error:
    Call oForm.Freeze(False)
   
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "PH_PY015_SAVE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub PH_PY015_Delete()
    '//선택된 자료 삭제
    
    Dim CLTCOD As String
    Dim MSTCOD As String
    Dim YEAR As String
    Dim FullName As String
    
    Dim i As Integer
    Dim cnt As Integer
    
    Dim sQry    As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY015_Delete_Error

    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    
    
    Call oForm.Freeze(True)
    
    CLTCOD = oForm.Items("CLTCOD").Specific.VALUE
    YEAR = oForm.Items("Year").Specific.VALUE
    MSTCOD = oForm.Items("MSTCOD").Specific.VALUE
    FullName = oForm.Items("FullName").Specific.VALUE
    
    sQry = " Select Count(*) From [@PH_PY015A] Where U_CLTCOD = '" & CLTCOD & "' And U_Year = '" & YEAR & "' And U_MSTCOD = '" & MSTCOD & "'"
    oRecordSet.DoQuery sQry
    
    cnt = oRecordSet.Fields(0).VALUE
    If cnt > 0 Then
    
        If Trim(YEAR) = "" Then
            Call MDC_Com.MDC_GF_Message("년도가 없습니다. 확인바랍니다..", "E")
            GoTo PH_PY015_Delete_Exit
        End If
        
        If Trim(CLTCOD) = "" Then
            Call MDC_Com.MDC_GF_Message("사업장이 없습니다. 확인바랍니다..", "E")
            GoTo PH_PY015_Delete_Exit
        End If
        If Trim(MSTCOD) = "" Then
            Call MDC_Com.MDC_GF_Message("사번이 없습니다. 확인바랍니다..", "E")
            GoTo PH_PY015_Delete_Exit
        End If
        
        
        
        
        If Sbo_Application.MessageBox(" 선택한사원('" & FullName & "')을 삭제하시겠습니까? ?", "2", "예", "아니오") = "1" Then
            sQry = "Delete From [@PH_PY015A] Where U_CLTCOD = '" & CLTCOD & "' AND  U_Year = '" & YEAR & "' And U_MSTCOD = '" & MSTCOD & "' "
            oRecordSet.DoQuery sQry
        End If
    End If
    
   
    Call oForm.Freeze(False)
    
        
    Call PH_PY015_MTX01
    
    Set oRecordSet = Nothing
    
    
    Exit Sub
PH_PY015_Delete_Exit:
    Set oRecordSet = Nothing
        
    Call oForm.Freeze(False)
    Exit Sub
PH_PY015_Delete_Error:
    Set oRecordSet = Nothing
   
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY015_Delete_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub PH_PY015_TitleSetting(iRow As Long)
    Dim i               As Long
    Dim j               As Long
    Dim sQry            As String
     
    Dim COLNAM(7)       As String

    Dim oColumn         As SAPbouiCOM.EditTextColumn
    Dim oComboCol       As SAPbouiCOM.ComboBoxColumn
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
     
    On Error GoTo Error_Message
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    oForm.Freeze True
    
    COLNAM(0) = "년도"
    COLNAM(1) = "부서"
    COLNAM(2) = "담당"
    COLNAM(3) = "사번"
    COLNAM(4) = "성명"
    COLNAM(5) = "직급"
    COLNAM(6) = "적치일수"
    COLNAM(7) = "지급일수"

    For i = 0 To UBound(COLNAM)
        oGrid1.Columns.Item(i).TitleObject.Caption = COLNAM(i)
        oGrid1.Columns.Item(i).Editable = False
    Next i
    
    oGrid1.Columns.Item(6).RightJustified = True
    oGrid1.Columns.Item(7).RightJustified = True
    
    oGrid1.AutoResizeColumns

    oForm.Freeze False

    Set oColumn = Nothing

    Exit Sub

Error_Message:
    oForm.Freeze False
    Set oColumn = Nothing
    Sbo_Application.SetStatusBarMessage "PH_PY015_TitleSetting Error : " & Space(10) & Err.Description, bmt_Short, True
End Sub

