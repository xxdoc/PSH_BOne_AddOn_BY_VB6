VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY679"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY679.cls
'//  Module         : 인사관리 > 근태관리
'//  Desc           : 개인별근태자료 집계(N.G.Y)
'//********************************************************************************
Option Explicit

Public oFormUniqueID    As String
Public oForm            As SAPbouiCOM.Form

Public oGrid1           As SAPbouiCOM.Grid
Public oDS_PH_PY679     As SAPbouiCOM.DataTable

Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long

Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY679.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY679_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY679"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE

        
    oForm.Freeze True
    Call PH_PY679_CreateItems
    Call PH_PY679_EnableMenus
    Call PH_PY679_SetDocument(oFromDocEntry01)
'    Call PH_PY679_FormResize
    
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

Private Function PH_PY679_CreateItems() As Boolean

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
    
    On Error GoTo PH_PY679_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Set oGrid1 = oForm.Items("Grid01").Specific
    
    oForm.DataSources.DataTables.Add ("PH_PY679")
    
'    oForm.DataSources.DataTables.Item("PH_PY679").Columns.Add "구분", ft_AlphaNumeric
'    oForm.DataSources.DataTables.Item("PH_PY679").Columns.Add "부서", ft_AlphaNumeric
'    oForm.DataSources.DataTables.Item("PH_PY679").Columns.Add "담당", ft_AlphaNumeric
'    oForm.DataSources.DataTables.Item("PH_PY679").Columns.Add "사번", ft_AlphaNumeric
'    oForm.DataSources.DataTables.Item("PH_PY679").Columns.Add "성명", ft_AlphaNumeric
'    oForm.DataSources.DataTables.Item("PH_PY679").Columns.Add "횟수", ft_AlphaNumeric
'    oForm.DataSources.DataTables.Item("PH_PY679").Columns.Add "교통비", ft_AlphaNumeric
        
    oGrid1.DataTable = oForm.DataSources.DataTables.Item("PH_PY679")
    Set oDS_PH_PY679 = oForm.DataSources.DataTables.Item("PH_PY679")
    '//----------------------------------------------------------------------------------------------
    '// 기본사항
    '//----------------------------------------------------------------------------------------------
    
    '//사업장
    Call oForm.DataSources.UserDataSources.Add("CLTCOD", dt_SHORT_TEXT, 10)
    Set oCombo = oForm.Items("CLTCOD").Specific
    oCombo.DataBind.SetBound True, "", "CLTCOD"
'    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
'    Call SetReDataCombo(oForm, sQry, oCombo)
'
'    CLTCOD = MDC_SetMod.Get_ReData("Branch", "USER_CODE", "OUSR", "'" & oCompany.UserName & "'")
'    oCombo.Select CLTCOD, psk_ByValue
    
    oForm.Items("CLTCOD").DisplayDesc = True
    
        
     '//기준일자
    Call oForm.DataSources.UserDataSources.Add("DocDateFr", dt_DATE, 10)
    oForm.Items("DocDateFr").Specific.DataBind.SetBound True, "", "DocDateFr"
    oForm.DataSources.UserDataSources.Item("DocDateFr").VALUE = Format(Now, "YYYYMMDD")
    
    Call oForm.DataSources.UserDataSources.Add("DocDateTo", dt_DATE, 10)
    oForm.Items("DocDateTo").Specific.DataBind.SetBound True, "", "DocDateTo"
    oForm.DataSources.UserDataSources.Item("DocDateTo").VALUE = Format(Now, "YYYYMMDD")
    
    
    '//출력구분
    Set oCombo = oForm.Items("Div").Specific
    oCombo.ValidValues.Add "1", "현재소속부서기준"
    oCombo.ValidValues.Add "2", "근무부서기준"
'    oCombo.ValidValues.Add "3", "직급별"
'    oCombo.ValidValues.Add "4", "판관제조별(월별)"
    oCombo.Select 0, psk_Index
    
    '//부서
    Call oForm.DataSources.UserDataSources.Add("TeamCode", dt_SHORT_TEXT, 10)
    Set oCombo = oForm.Items("TeamCode").Specific
    oCombo.DataBind.SetBound True, "", "TeamCode"
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '1' AND U_UseYN= 'Y' AND U_Char2 = '" & oForm.Items.Item("CLTCOD").Specific.VALUE & "'"
    Call SetReDataCombo(oForm, sQry, oCombo, "Y")
    oForm.Items("TeamCode").DisplayDesc = True
    
    '//담당
    Call oForm.DataSources.UserDataSources.Add("RspCode", dt_SHORT_TEXT, 10)
    Set oCombo = oForm.Items("RspCode").Specific
    oCombo.DataBind.SetBound True, "", "RspCode"
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '2' AND U_UseYN= 'Y' AND U_Char2 = '" & oForm.Items.Item("CLTCOD").Specific.VALUE & "'"
    Call SetReDataCombo(oForm, sQry, oCombo, "Y")
    oForm.Items("RspCode").DisplayDesc = True
    
     '//커서를 첫번째 ITEM으로 지정
    oForm.ActiveItem = "CLTCOD"
    

    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
    
PH_PY679_CreateItems_Error:
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY679_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY679_EnableMenus()
    
    On Error GoTo PH_PY679_EnableMenus_Error

    Call oForm.EnableMenu("1283", True)     '//제거
    Call oForm.EnableMenu("1284", False)    '//취소
    Call oForm.EnableMenu("1293", True)     '//행삭제
    
    Exit Sub
    
PH_PY679_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY679_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY679_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY679_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY679_FormItemEnabled
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY679_FormItemEnabled
        oForm.Items("Code").Specific.VALUE = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY679_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY679_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY679_FormItemEnabled()
    Dim oCombo      As SAPbouiCOM.ComboBox
    
    On Error GoTo PH_PY679_FormItemEnabled_Error
    
    
    Call oForm.Freeze(True)
    If (oForm.Mode = fm_ADD_MODE) Then
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        
        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", False)    '//문서추가
        
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
    Call oForm.Freeze(False)
    Exit Sub
    
PH_PY679_FormItemEnabled_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY679_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
                If pval.ItemUID = "Btn_Serch" Then
                    If PH_PY679_DataValidCheck = True Then
                        Call PH_PY679_DataFind
                    Else
                        BubbleEvent = False
                    End If
                End If
                If pval.ItemUID = "Btn_Prt" Then
                    If PH_PY679_DataValidCheck = True Then
                        Call PH_PY679_Print_Report01
                    Else
                        BubbleEvent = False
                    End If
                End If
            ElseIf pval.BeforeAction = False Then
'                If oForm.Mode = fm_ADD_MODE Then
'                    If pval.ActionSuccess = True Then
'                        Call PH_PY679_FormItemEnabled
'                    End If
'                ElseIf oForm.Mode = fm_UPDATE_MODE Then
'                    If pval.ActionSuccess = True Then
'                        Call PH_PY679_FormItemEnabled
'                    End If
'                ElseIf oForm.Mode = fm_OK_MODE Then
'                    If pval.ActionSuccess = True Then
'                        Call PH_PY679_FormItemEnabled
'                    End If
'                End If

            End If
     
'----------------------------------------------------------
        Case et_KEY_DOWN: '//2

'----------------------------------------------------------
        Case et_GOT_FOCUS: '//3
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
'----------------------------------------------------------
        Case et_LOST_FOCUS: '//4

'----------------------------------------------------------
        Case et_COMBO_SELECT: '//5
            Call oForm.Freeze(True)
            If pval.BeforeAction = True Then
                
            ElseIf pval.BeforeAction = False Then
                If pval.ItemChanged = True Then
                    Select Case pval.ItemUID
                           '//사업장이 바뀌면 부서와 담당 재설정
                           Case "CLTCOD"
                                '//부서
                                Set oCombo = oForm.Items("TeamCode").Specific
                                If oCombo.ValidValues.Count > 0 Then              '//삭제
                                   For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                                           oCombo.ValidValues.Remove i, psk_Index
                                   Next i
                                End If
                                '//현재 사업장으로 다시 Qry
                                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '1' AND U_UseYN= 'Y' AND U_Char2 = '" & oForm.Items.Item("CLTCOD").Specific.VALUE & "'"
                                Call SetReDataCombo(oForm, sQry, oCombo, "Y")
                                
                                '//담당
                                Set oCombo = oForm.Items("RspCode").Specific
                                If oCombo.ValidValues.Count > 0 Then              '//삭제
                                   For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                                           oCombo.ValidValues.Remove i, psk_Index
                                   Next i
                                End If
                                '//현재 사업장으로 다시 Qry
                                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '2' AND U_UseYN= 'Y' AND U_Char2 = '" & oForm.Items.Item("CLTCOD").Specific.VALUE & "'"
                                Call SetReDataCombo(oForm, sQry, oCombo, "Y")
                                
                           '//부서가 바뀌면 담당 재설정
                           Case "TeamCode"
                                '//담당은 그 부서의 담당만 표시
                                Set oCombo = oForm.Items("RspCode").Specific
                                If oCombo.ValidValues.Count > 0 Then              '//삭제
                                   For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                                           oCombo.ValidValues.Remove i, psk_Index
                                   Next i
                                End If
                                '//현재 사업장으로 다시 Qry
                                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '2' AND U_UseYN= 'Y' AND U_Char1 = '" & oForm.Items.Item("TeamCode").Specific.VALUE & "' AND U_Char2 = '" & oForm.Items.Item("CLTCOD").Specific.VALUE & "'"
                                Call SetReDataCombo(oForm, sQry, oCombo, "Y")
                           
                                
                           
                    End Select
                End If
            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_CLICK: '//6
            If pval.BeforeAction = True Then
                Select Case pval.ItemUID
                Case "Grid01"
                    If pval.Row > 0 Then
'                        Call oGrid1.SelectRow(pval.Row, True, False)
                        
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
'                    Select Case pval.ItemUID
'                        Case "MSTCOD"
'                            '//사원명 찿아서 화면 표시 하기
'                            sQry = "SELECT U_FullName FROM [@PH_PY001A] WHERE Code =  '" & Trim(oForm.Items("MSTCOD").Specific.VALUE) & "'"
'                            oRecordSet.DoQuery sQry
'                            oForm.Items("MSTNAME").Specific.String = Trim(oRecordSet.Fields("U_FullName").VALUE)
'
'                    End Select
                End If
            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_MATRIX_LOAD: '//11
            If pval.BeforeAction = True Then
            ElseIf pval.BeforeAction = False Then
                Call PH_PY679_FormItemEnabled
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
                Set oDS_PH_PY679 = Nothing
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
                Call PH_PY679_FormItemEnabled

            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//문서찾기
                Call PH_PY679_FormItemEnabled
            Case "1282":    '//문서추가
                Call PH_PY679_FormItemEnabled

            Case "1288", "1289", "1290", "1291":
                Call PH_PY679_FormItemEnabled
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
    Exit Sub
Raise_RightClickEvent_Error:

    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY679_FormClear()
On Error GoTo PH_PY679_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY679'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PH_PY679_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY679_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY679_DataValidCheck() As Boolean
On Error GoTo PH_PY679_DataValidCheck_Error
    PH_PY679_DataValidCheck = False
    Dim i       As Long
    Dim sQry    As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    If oForm.Items("CLTCOD").Specific.VALUE = "" Then
        Sbo_Application.SetStatusBarMessage "사업장은 필수입니다.", bmt_Short, True
        oForm.Items("CLTCOD").CLICK ct_Regular
        PH_PY679_DataValidCheck = False
        Exit Function
    End If
    
    PH_PY679_DataValidCheck = True
    
        
    Exit Function
    
    Set oRecordSet = Nothing
    
    
PH_PY679_DataValidCheck_Error:
    Set oRecordSet = Nothing
    PH_PY679_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY679_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Function PH_PY679_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY679_Validate_Error
    PH_PY679_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY679A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", bmt_Short, True
        PH_PY679_Validate = False
        GoTo PH_PY679_Validate_Exit
    End If
'
    If ValidateType = "수정" Then

    ElseIf ValidateType = "행삭제" Then

    ElseIf ValidateType = "취소" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY679_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY679_Validate_Error:
    PH_PY679_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY679_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Private Sub PH_PY679_DataFind()
    Dim i           As Long
    Dim iRow        As Long
    Dim sQry        As String
    Dim CLTCOD      As String
    Dim DocDateFr   As String
    Dim DocDateTo   As String
    Dim Div         As String
    Dim TeamCode    As String
    Dim RspCode     As String
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY679_DataFind_Error
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE)
    DocDateFr = Trim(oForm.Items("DocDateFr").Specific.VALUE)
    DocDateTo = Trim(oForm.Items("DocDateTo").Specific.VALUE)
    Div = Trim(oForm.Items("Div").Specific.VALUE)
    TeamCode = Trim(oForm.Items("TeamCode").Specific.VALUE)
    RspCode = Trim(oForm.Items("RspCode").Specific.VALUE)
    
    sQry = "Exec PH_PY679_01 '" & CLTCOD & "','" & DocDateFr & "','" & DocDateTo & "', '" & TeamCode & "', '" & RspCode & "', '" & Div & "'"
'    If Div = "1" Then
'       sQry = "Exec PH_PY679_01 '" & CLTCOD & "','" & DocDateFr & "','" & DocDateTo & "'"
'    ElseIf Div = "2" Then
'       sQry = "Exec PH_PY679_02 '" & CLTCOD & "','" & DocDateFr & "','" & DocDateTo & "'"
'    ElseIf Div = "3" Then
'       sQry = "Exec PH_PY679_03 '" & CLTCOD & "','" & DocDateFr & "','" & DocDateTo & "'"
'    ElseIf Div = "4" Then
'       sQry = "Exec PH_PY679_04 '" & CLTCOD & "','" & DocDateFr & "','" & DocDateTo & "'"
'    End If
    
    Call oDS_PH_PY679.ExecuteQuery(sQry)
        
    iRow = oForm.DataSources.DataTables.Item(0).Rows.Count
    
   ' Call PH_PY679_TitleSetting(iRow)
    
    Set oRecordSet = Nothing
    Exit Sub
    
PH_PY679_DataFind_Error:
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "PH_PY679_DataFind_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PH_PY679_DataSave() As Boolean
    Dim i           As Long
    Dim ShiftDat    As String
    Dim sQry        As String
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    On Error GoTo PH_PY679_DataSave_Error
    
    PH_PY679_DataSave = False

    If oForm.DataSources.DataTables.Item(0).Rows.Count > 0 Then
        For i = 0 To oForm.DataSources.DataTables.Item(0).Rows.Count - 1
'            oDS_PH_PY679.Columns.Item("Code").Cells(i).Value
            
            sQry = " UPDATE [@PH_PY001A] SET U_Activity = '" & oDS_PH_PY679.Columns.Item("Activity").Cells(i).VALUE & "'"
            sQry = sQry & " WHERE Code = '" & oDS_PH_PY679.Columns.Item("MSTCOD").Cells(i).VALUE & "'"
            oRecordSet.DoQuery sQry
            
            
            
        Next i
        Sbo_Application.SetStatusBarMessage "기본업무가 변경되었습니다.", bmt_Short, False
        PH_PY679_DataSave = True
    Else
        Sbo_Application.SetStatusBarMessage "데이터가 존재하지 않습니다.", bmt_Short, True
        PH_PY679_DataSave = False
    End If
    
    Set oRecordSet = Nothing
    Exit Function
    
PH_PY679_DataSave_Error:
    Set oRecordSet = Nothing
    PH_PY679_DataSave = False
    Sbo_Application.SetStatusBarMessage "PH_PY679_DataSave_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

'Private Sub PH_PY679_TitleSetting(iRow As Long)
'    Dim i               As Long
'    Dim j               As Long
'    Dim sQry            As String
'
'    Dim COLNAM(6)       As String
'
'    Dim oColumn         As SAPbouiCOM.EditTextColumn
'    Dim oComboCol       As SAPbouiCOM.ComboBoxColumn
'
'    Dim oRecordSet  As SAPbobsCOM.Recordset
'
'    On Error GoTo Error_Message
'
'    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
'
'    oForm.Freeze True
'
'    COLNAM(0) = "구분"
'    COLNAM(1) = "부서"
'    COLNAM(2) = "담당"
'    COLNAM(3) = "사번"
'    COLNAM(4) = "성명"
'    COLNAM(5) = "횟수"
'    COLNAM(6) = "교통비"
'
'    For i = 0 To UBound(COLNAM)
'        oGrid1.Columns.Item(i).TitleObject.Caption = COLNAM(i)
'        If i >= 0 And i < UBound(COLNAM) Then
'            oGrid1.Columns.Item(i).Editable = False
''        ElseIf i = UBound(COLNAM) Then
''            oGrid1.Columns.Item(i).Editable = True
''            oGrid1.Columns.Item(i).Type = gct_ComboBox
''            Set oComboCol = oGrid1.Columns.Item("Activity")
''
''            sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
''            sQry = sQry & " WHERE U_Char2 = '" & oForm.Items("CLTCOD").Specific.Value & "' And Code = 'P127' AND U_UseYN= 'Y' Order by U_Seq"
''            oRecordSet.DoQuery sQry
''            If oRecordSet.RecordCount > 0 Then
''                    oComboCol.ValidValues.Add "", ""
''                For j = 0 To oRecordSet.RecordCount - 1
''                    oComboCol.ValidValues.Add oRecordSet.Fields(0).Value, oRecordSet.Fields(1).Value
''                    oRecordSet.MoveNext
''                Next j
'''                oComboCol.Select 0, psk_Index
''            End If
''
''
''            oComboCol.DisplayType = cdt_Description
'        End If
'
'    Next i
'
'    oGrid1.AutoResizeColumns
'
'    oForm.Freeze False
'
'    Set oColumn = Nothing
'
'    Exit Sub

'Error_Message:
'    oForm.Freeze False
'    Set oColumn = Nothing
'    Sbo_Application.SetStatusBarMessage "PH_PY679_TitleSetting Error : " & Space(10) & Err.Description, bmt_Short, True
'End Sub

Private Sub PH_PY679_Print_Report01()

'    Dim DocNum          As String
'    Dim ErrNum          As Integer
'    Dim WinTitle        As String
'    Dim ReportName      As String
'    Dim sQry            As String
'    Dim oRecordSet      As SAPbobsCOM.Recordset
'
'    On Error GoTo PH_PY679_Print_Report01_Error
'
'    Dim CLTCOD          As String
'    Dim DocDateFr       As String
'    Dim DocDateTo       As String
'    Dim MSTCOD          As String
'    Dim Div             As String
'
'    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
'
'
'     '/ ODBC 연결 체크
'    If ConnectODBC = False Then
'        GoTo PH_PY679_Print_Report01_Error
'    End If
'
'
'    '//인자 MOVE , Trim 시키기..
'    CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE)
'    DocDateFr = Trim(oForm.Items("DocDateFr").Specific.VALUE)
'    DocDateTo = Trim(oForm.Items("DocDateTo").Specific.VALUE)
'    MSTCOD = Trim(oForm.Items("MSTCOD").Specific.VALUE)
'    Div = Trim(oForm.Items("Div").Specific.VALUE)
'
'
'    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
'    If Div = "1" Then
'       WinTitle = "[PH_PY679] 개인별교통비지급대장"
'       ReportName = "PH_PY679_01.rpt"
'    ElseIf Div = "2" Then
'       WinTitle = "[PH_PY679] 개인별근태월보(교통비확인용)"
'       ReportName = "PH_PY679_02.rpt"
'    ElseIf Div = "3" Then
'       WinTitle = "[PH_PY679] 통근버스휴일운행현황"
'       ReportName = "PH_PY679_03.rpt"
'    End If
'
'
'    '/ Formula 수식필드
'    ReDim gRpt_Formula(3)
'    ReDim gRpt_Formula_Value(3)
'
'    gRpt_Formula(1) = "CLTCOD"
'    sQry = "SELECT U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y' AND U_Code = '" & CLTCOD & "'"
'    Call oRecordSet.DoQuery(sQry)
'    gRpt_Formula_Value(1) = oRecordSet.Fields(0).VALUE
'
'    gRpt_Formula(2) = "DocDateFr"
'    gRpt_Formula_Value(2) = Format(DocDateFr, "####-##-##")
'
'    gRpt_Formula(3) = "DocDateTo"
'    gRpt_Formula_Value(3) = Format(DocDateTo, "####-##-##")
'
'
'    '/ SubReport
'    ReDim gRpt_SRptSqry(1)
'    ReDim gRpt_SRptName(1)
'
'    ReDim gRpt_SFormula(1, 1)
'    ReDim gRpt_SFormula_Value(1, 1)
'
'
'    '/ Procedure 실행"
'    If Div = "1" Then
'       sQry = "EXEC [PH_PY679_01] '" & CLTCOD & "', '" & DocDateFr & "',  '" & DocDateTo & "',  '" & MSTCOD & "'"
'    ElseIf Div = "2" Then
'       sQry = "EXEC [PH_PY679_02] '" & CLTCOD & "', '" & DocDateFr & "',  '" & DocDateTo & "',  '" & MSTCOD & "'"
'    ElseIf Div = "3" Then
'       sQry = "EXEC [PH_PY679_03] '" & CLTCOD & "', '" & DocDateFr & "',  '" & DocDateTo & "'"
'    End If
'
'    oRecordSet.DoQuery sQry
'    If oRecordSet.RecordCount = 0 Then
'        ErrNum = 1
'        GoTo PH_PY679_Print_Report01_Error
'    End If
'
'    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "1", "Y", "V", , 1) = False Then
'        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
'    End If
'
'    Set oRecordSet = Nothing
'    Exit Sub
    
'PH_PY679_Print_Report01_Error:
'    If ErrNum = 1 Then
'        Set oRecordSet = Nothing
'        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
'    Else
'    Set oRecordSet = Nothing
'    Sbo_Application.SetStatusBarMessage "PH_PY679_Print_Report01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
'    End If
    
End Sub






