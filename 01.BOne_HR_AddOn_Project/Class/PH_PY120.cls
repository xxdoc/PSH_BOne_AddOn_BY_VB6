VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY120"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY120.cls
'//  Module         : 급여관리 > 급상여소급 집계처리
'//  Desc           : 급상여소급 집계로 급여파일로 생성
'//********************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form

'// 그리드 사용시
Public oGrid1           As SAPbouiCOM.Grid
Public oDS_PH_PY120     As SAPbouiCOM.DataTable


Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long

Private sCLTCOD     As String
Private sYM         As String
Private sJOBTYP     As String
Private sJOBGBN     As String
Private sPAYSEL     As String
Private sMSTCOD     As String
Private sTeamCode   As String
Private sRspCode    As String
Private sClose      As String

Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY120.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY120_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY120"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
'    oForm.DataBrowser.BrowseBy = "Code"
        
    oForm.Freeze True
    Call PH_PY120_CreateItems
    Call PH_PY120_EnableMenus
    Call PH_PY120_SetDocument(oFromDocEntry01)
'    Call PH_PY120_FormResize
    
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

Private Function PH_PY120_CreateItems() As Boolean

    Dim sQry        As String
    Dim i           As Long
    
    Dim oCheck      As SAPbouiCOM.CheckBox
    Dim oEdit       As SAPbouiCOM.EditText
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns    As SAPbouiCOM.Columns
    Dim optBtn      As SAPbouiCOM.OptionBtn
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY120_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    
    '//----------------------------------------------------------------------------------------------
    '// 데이터셋정의
    '//----------------------------------------------------------------------------------------------

    '//테이블이 없는경우 데이터셋(Grid)
    oForm.DataSources.DataTables.Add ("PH_PY120")
    oForm.DataSources.DataTables.Item("PH_PY120").Columns.Add "부서", ft_AlphaNumeric
    oForm.DataSources.DataTables.Item("PH_PY120").Columns.Add "담당", ft_AlphaNumeric
    oForm.DataSources.DataTables.Item("PH_PY120").Columns.Add "사번", ft_AlphaNumeric
    oForm.DataSources.DataTables.Item("PH_PY120").Columns.Add "성명", ft_AlphaNumeric
    oForm.DataSources.DataTables.Item("PH_PY120").Columns.Add "지급구분", ft_AlphaNumeric
    oForm.DataSources.DataTables.Item("PH_PY120").Columns.Add "총지급액", ft_AlphaNumeric
    oForm.DataSources.DataTables.Item("PH_PY120").Columns.Add "총공제액", ft_AlphaNumeric
    oForm.DataSources.DataTables.Item("PH_PY120").Columns.Add "실지급액", ft_AlphaNumeric

    Set oGrid1 = oForm.Items("Grid1").Specific

    oGrid1.DataTable = oForm.DataSources.DataTables.Item("PH_PY120")
    Set oDS_PH_PY120 = oForm.DataSources.DataTables.Item("PH_PY120")
    
    
    '//----------------------------------------------------------------------------------------------
    '// 아이템 설정
    '//----------------------------------------------------------------------------------------------

'    '//UserDataSources
'    Call oForm.DataSources.UserDataSources.Add("     ", dt_SHORT_TEXT, 10)
'    Set oCombo = oForm.Items("    ").Specific
'    oCombo.DataBind.SetBound True, "", "    "
'    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
'    Call SetReDataCombo(oForm, sQry, oCombo)
'    oForm.Items("CLTCOD").DisplayDesc = True
    
    
    '//사업장
    Call oForm.DataSources.UserDataSources.Add("CLTCOD", dt_SHORT_TEXT, 1)
    Set oCombo = oForm.Items("CLTCOD").Specific
    oCombo.DataBind.SetBound True, "", "CLTCOD"
    Call CLTCOD_Select(oForm, "CLTCOD")
    oForm.Items("CLTCOD").DisplayDesc = True
    
    '// 귀속년월
    Call oForm.DataSources.UserDataSources.Add("YM", dt_SHORT_TEXT, 6)
    Set oEdit = oForm.Items("YM").Specific
    oEdit.DataBind.SetBound True, "", "YM"
    oForm.DataSources.UserDataSources.Item("YM").ValueEx = Format(Date, "YYYYMM")
'    Call oDS_PH_PY120.setValue("U_YM", 0, Format(Date, "YYYYMM"))
    
    '//대상기간
    Call oForm.DataSources.UserDataSources.Add("YMFrom", dt_SHORT_TEXT, 6)
    Set oEdit = oForm.Items("YMFrom").Specific
    oEdit.DataBind.SetBound True, "", "YMFrom"
    
    Call oForm.DataSources.UserDataSources.Add("YMTo", dt_SHORT_TEXT, 6)
    Set oEdit = oForm.Items("YMTo").Specific
    oEdit.DataBind.SetBound True, "", "YMTo"
    
    
'    Call oForm.DataSources.UserDataSources.Add("JIGBIL", dt_DATE, 10)
'    oForm.Items("JIGBIL").Specific.DataBind.SetBound True, "", "JIGBIL"
    
    '//지급일자
    Call oForm.DataSources.UserDataSources.Add("JIGBIL", dt_DATE, 8)
    Set oEdit = oForm.Items("JIGBIL").Specific
    oEdit.DataBind.SetBound True, "", "JIGBIL"
    
    
    
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
    
PH_PY120_CreateItems_Error:
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY120_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY120_EnableMenus()
    
    On Error GoTo PH_PY120_EnableMenus_Error

    Call oForm.EnableMenu("1283", True)     '//제거
    Call oForm.EnableMenu("1284", False)    '//취소
    Call oForm.EnableMenu("1293", True)     '//행삭제
    
    Exit Sub
    
PH_PY120_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY120_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY120_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY120_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY120_FormItemEnabled
        Call PH_PY120_AddMatrixRow
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY120_FormItemEnabled
        oForm.Items("Code").Specific.VALUE = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY120_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY120_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY120_FormItemEnabled()
    Dim oCombo      As SAPbouiCOM.ComboBox
    
    On Error GoTo PH_PY120_FormItemEnabled_Error
    
    
    Call oForm.Freeze(True)
    If (oForm.Mode = fm_ADD_MODE) Then

        Call oForm.EnableMenu("1281", False)     '//문서찾기
        Call oForm.EnableMenu("1282", False)    '//문서추가
        
    ElseIf (oForm.Mode = fm_FIND_MODE) Then

        Call oForm.EnableMenu("1281", False)    '//문서찾기
        Call oForm.EnableMenu("1282", False)     '//문서추가
    ElseIf (oForm.Mode = fm_OK_MODE) Then

        
        Call oForm.EnableMenu("1281", False)     '//문서찾기
        Call oForm.EnableMenu("1282", False)     '//문서추가

    End If
    Call oForm.Freeze(False)
    Exit Sub
    
PH_PY120_FormItemEnabled_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY120_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Public Sub Raise_FormItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
    Dim sQry        As String
    Dim i           As Long
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo Raise_FormItemEvent_Error
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Select Case pval.EventType
        Case et_ITEM_PRESSED: '//1
            
            If pval.BeforeAction = True Then
                
            ElseIf pval.BeforeAction = False Then
                If pval.ItemUID = "Btn1" Then
                    If Sbo_Application.MessageBox("월별로 계산된 소급자료를 해당월, 지급일자로 소급급여처리를 진행 하시겠습니까?", 2, "Yes", "No") = 2 Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                    Call PH_PY120_DataChange
                    Call PH_PY120_DataSearch
                End If
                If pval.ItemUID = "Btn_Search" Then
                    Call PH_PY120_DataSearch
                End If
                
            End If
     
'----------------------------------------------------------
        Case et_KEY_DOWN: '//2

'----------------------------------------------------------
        Case et_GOT_FOCUS: '//3
            Select Case pval.ItemUID
            Case "Mat1", "Grid1"
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
                    If pval.ItemUID = "CLTCOD" Then
                    
                    End If
                End If
            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_CLICK: '//6

'----------------------------------------------------------
        Case et_DOUBLE_CLICK: '//7

'----------------------------------------------------------
        Case et_MATRIX_LINK_PRESSED '//8

'----------------------------------------------------------
        Case et_MATRIX_COLLAPSE_PRESSED '//9
'----------------------------------------------------------
        Case et_VALIDATE: '//10
            Call oForm.Freeze(True)
            If pval.BeforeAction = True Then

            ElseIf pval.BeforeAction = False Then
                If pval.ItemChanged = True Then
                    If pval.ItemUID = "MSTCOD" Then
                        
                    End If
                End If
            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_MATRIX_LOAD: '//11
'            If pval.BeforeAction = True Then
'            ElseIf pval.BeforeAction = False Then
'                oMat1.LoadFromDataSource
'
'                Call PH_PY120_FormItemEnabled
'                Call PH_PY120_AddMatrixRow
'
'            End If
'----------------------------------------------------------
        Case et_DATASOURCE_LOAD '//12
           
'----------------------------------------------------------
        Case et_FORM_LOAD: '//16
            
'----------------------------------------------------------
        Case et_FORM_UNLOAD: '//17
            'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
            '컬렉션에서 삭제및 모든 메모리 제거
            'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
            If pval.BeforeAction = True Then
            ElseIf pval.BeforeAction = False Then
                RemoveForms oFormUniqueID
                Set oForm = Nothing
                Set oDS_PH_PY120 = Nothing

                Set oGrid1 = Nothing
                
            End If
'----------------------------------------------------------
        Case et_FORM_ACTIVATE: '//18
            
'----------------------------------------------------------
        Case et_FORM_DEACTIVATE: '//19
            
'----------------------------------------------------------
        Case et_FORM_CLOSE '//20
            
'----------------------------------------------------------
        Case et_FORM_RESIZE '//21
'            If pval.BeforeAction = True Then
'
'            ElseIf pval.BeforeAction = False Then
'
'            End If
'----------------------------------------------------------
        Case et_FORM_KEY_DOWN '//22
            
'----------------------------------------------------------
        Case et_FORM_MENU_HILIGHT '//23
            
'----------------------------------------------------------
        Case et_CHOOSE_FROM_LIST '//27
'            If pval.BeforeAction = True Then
'
'            ElseIf pval.Before_Action = False Then
'
'            End If
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
    Sbo_Application.SetStatusBarMessage "Raise_ItemEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
        End Select
    ElseIf (pval.BeforeAction = False) Then
        Select Case pval.MenuUID
            Case "1283":
                oForm.Mode = fm_ADD_MODE
                Call PH_PY120_FormItemEnabled
                Call PH_PY120_AddMatrixRow
            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//문서찾기
                Call PH_PY120_FormItemEnabled
                Call PH_PY120_AddMatrixRow
                oForm.Items("Code").CLICK ct_Regular
            Case "1282":    '//문서추가
                Call PH_PY120_FormItemEnabled
                Call PH_PY120_AddMatrixRow
            Case "1288", "1289", "1290", "1291":
                Call PH_PY120_FormItemEnabled
            Case "1293" '// 행삭제

        End Select
    End If
    Call oForm.Freeze(False)
    Exit Sub
Raise_FormMenuEvent_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_FormMenuEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub Raise_FormDataEvent(ByRef FormUID As String, ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
    
    On Error GoTo Raise_FormDataEvent_Error
    
    If (BusinessObjectInfo.BeforeAction = True) Then
        Select Case BusinessObjectInfo.EventType
            Case et_FORM_DATA_LOAD:     '//33
            Case et_FORM_DATA_ADD:      '//34
            Case et_FORM_DATA_UPDATE:   '//35
            Case et_FORM_DATA_DELETE:   '//36
        End Select
    ElseIf (BusinessObjectInfo.BeforeAction = False) Then
        Select Case BusinessObjectInfo.EventType
            Case et_FORM_DATA_LOAD:     '//33
            Case et_FORM_DATA_ADD:      '//34
            Case et_FORM_DATA_UPDATE:   '//35
            Case et_FORM_DATA_DELETE:   '//36
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
    Select Case pval.ItemUID
    Case "Mat1"
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

Sub PH_PY120_AddMatrixRow()
    Dim oRow    As Long
    
    On Error GoTo PH_PY120_AddMatrixRow_Error
    
    Call oForm.Freeze(True)
   
    Call oForm.Freeze(False)
    Exit Sub
PH_PY120_AddMatrixRow_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY120_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY120_FormClear()
On Error GoTo PH_PY120_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY120'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PH_PY120_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY120_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY120_DataValidCheck() As Boolean
On Error GoTo PH_PY120_DataValidCheck_Error
    PH_PY120_DataValidCheck = False
    Dim i       As Long
    Dim sQry    As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '//----------------------------------------------------------------------------------
    '//필수 체크
    '//----------------------------------------------------------------------------------
    
    PH_PY120_DataValidCheck = True
    
    Exit Function
    
    Set oRecordSet = Nothing
    
    
PH_PY120_DataValidCheck_Error:
    Set oRecordSet = Nothing
    PH_PY120_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY120_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Function PH_PY120_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY120_Validate_Error
    PH_PY120_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY120A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", bmt_Short, True
        PH_PY120_Validate = False
        GoTo PH_PY120_Validate_Exit
    End If
'
    If ValidateType = "수정" Then

    ElseIf ValidateType = "행삭제" Then

    ElseIf ValidateType = "취소" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY120_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY120_Validate_Error:
    PH_PY120_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY120_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY120_DataSearch()
    Dim i           As Long
    Dim iRow        As Long
    Dim sQry        As String
    
    Dim CLTCOD  As String
    Dim YM  As String
    Dim YMFrom  As String
    Dim YMTo  As String
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY120_DataSearch_Error

    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    oForm.Freeze (True)
    
    CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE)
    YM = Trim(oForm.Items("YM").Specific.VALUE)
    YMFrom = Trim(oForm.Items("YMFrom").Specific.VALUE)
    YMTo = Trim(oForm.Items("YMTo").Specific.VALUE)
    
    sQry = "Exec PH_PY120_01 '" & CLTCOD & "','" & YMFrom & "','" & YMTo & "'"
        
    Call oDS_PH_PY120.ExecuteQuery(sQry)
    
    Debug.Print oDS_PH_PY120.Rows.Count
    iRow = oForm.DataSources.DataTables.Item(0).Rows.Count
    
    Call PH_PY120_TitleSetting(iRow)
    
    Set oRecordSet = Nothing
    oForm.Freeze (False)
    Exit Sub
    
PH_PY120_DataSearch_Error:
    Set oRecordSet = Nothing
    oForm.Freeze (False)
    Sbo_Application.SetStatusBarMessage "PH_PY004_DataFind_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub PH_PY120_DataChange()
    Dim i           As Long
    Dim iRow        As Long
    Dim sQry        As String
    
    Dim CLTCOD  As String
    Dim YM      As String
    Dim JIGBIL  As String
    Dim YMFrom  As String
    Dim YMTo    As String
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY120_DataChange_Error

    CLTCOD = oForm.Items("CLTCOD").Specific.VALUE
    YM = oForm.Items("YM").Specific.VALUE
    JIGBIL = oForm.Items("JIGBIL").Specific.VALUE
    YMFrom = oForm.Items("YMFrom").Specific.VALUE
    YMTo = oForm.Items("YMTo").Specific.VALUE
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    oForm.Freeze (True)
    
    sQry = "Exec PH_PY111_SOGUBF '" & CLTCOD & "','" & YM & "','" & JIGBIL & "','" & YMFrom & "','" & YMTo & "'"
    oRecordSet.DoQuery (sQry)
    
    Sbo_Application.StatusBar.SetText "소급집계처리가 적용 되었습니다. 급여대장조회 확인 바랍니다.", bmt_Short, smt_Success
    
    Set oRecordSet = Nothing
    oForm.Freeze (False)
    Exit Sub
    
PH_PY120_DataChange_Error:
    Set oRecordSet = Nothing
    oForm.Freeze (False)
    Sbo_Application.SetStatusBarMessage "PH_PY004_DataFind_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY120_TitleSetting(iRow As Long)
    Dim i               As Long
    Dim j               As Long
    Dim sQry            As String
     
    Dim COLNAM(8)       As String

    Dim oColumn         As SAPbouiCOM.EditTextColumn
    Dim oComboCol       As SAPbouiCOM.ComboBoxColumn
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
     
    On Error GoTo Error_Message
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    oForm.Freeze True
    
    COLNAM(0) = "부서"
    COLNAM(1) = "담당"
    COLNAM(2) = "사번"
    COLNAM(3) = "성명"
    COLNAM(4) = "지급구분"
    COLNAM(5) = "총지급액"
    COLNAM(6) = "총공제액"
    COLNAM(7) = "실지급액"
    
    For i = 0 To UBound(COLNAM) - 1
        oGrid1.Columns.Item(i).TitleObject.Caption = COLNAM(i)
        oGrid1.Columns.Item(i).Editable = False
        
    Next i
    
    oGrid1.AutoResizeColumns

    oForm.Freeze False

    Set oColumn = Nothing

    Exit Sub

Error_Message:
    oForm.Freeze False
    Set oColumn = Nothing
    Sbo_Application.SetStatusBarMessage "PH_PY120_TitleSetting Error : " & Space(10) & Err.Description, bmt_Short, True
End Sub


