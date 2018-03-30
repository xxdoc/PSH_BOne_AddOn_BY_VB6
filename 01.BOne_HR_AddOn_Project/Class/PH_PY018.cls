VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY018"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY018.cls
'//  Module         : 인사관리 > 근태관리
'//  Desc           : 휴일근무체크(연봉제)
'//********************************************************************************
Option Explicit

Public oFormUniqueID    As String
Public oForm            As SAPbouiCOM.Form

Public oGrid1           As SAPbouiCOM.Grid

Public oDS_PH_PY018     As SAPbouiCOM.DataTable

Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long

Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY018.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY018_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY018"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE

        
    oForm.Freeze True
    Call PH_PY018_CreateItems
    Call PH_PY018_EnableMenus
    Call PH_PY018_SetDocument(oFromDocEntry01)
'    Call PH_PY018_FormResize
    
        '기간(월)
    oForm.Items("DocDateF").Specific.VALUE = Format(Now, "YYYYMM01")
    oForm.Items("DocDateT").Specific.VALUE = Format(Now, "YYYYMMDD")
    
    
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

Private Function PH_PY018_CreateItems() As Boolean

    Dim sQry        As String
    Dim i           As Long
    
    Dim oCheck      As SAPbouiCOM.CheckBox
    Dim oEdit       As SAPbouiCOM.EditText
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns    As SAPbouiCOM.Columns
    Dim optBtn      As SAPbouiCOM.OptionBtn
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY018_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Set oGrid1 = oForm.Items("Grid01").Specific
    oGrid1.SelectionMode = ms_Auto
    
    oForm.DataSources.DataTables.Add ("PH_PY018")
    
    oForm.DataSources.DataTables.Item("PH_PY018").Columns.Add "일자", ft_Date
    oForm.DataSources.DataTables.Item("PH_PY018").Columns.Add "사번", ft_AlphaNumeric
    oForm.DataSources.DataTables.Item("PH_PY018").Columns.Add "성명", ft_AlphaNumeric
    oForm.DataSources.DataTables.Item("PH_PY018").Columns.Add "요일", ft_AlphaNumeric
    oForm.DataSources.DataTables.Item("PH_PY018").Columns.Add "휴일근무(Y/N)", ft_AlphaNumeric
        
    oGrid1.DataTable = oForm.DataSources.DataTables.Item("PH_PY018")
    Set oDS_PH_PY018 = oForm.DataSources.DataTables.Item("PH_PY018")
    '//----------------------------------------------------------------------------------------------
    '// 기본사항
    '//----------------------------------------------------------------------------------------------
    
    
    '//사업장
'    Dim CLTCOD As String
'    CLTCOD = MDC_SetMod.Get_ReData("Branch", "USER_CODE", "OUSR", "'" & oCompany.UserName & "'")
    
'    Call oForm.DataSources.UserDataSources.Add("CLTCOD", dt_SHORT_TEXT, 10)
'    Set oCombo = oForm.Items("CLTCOD").Specific
'    oCombo.DataBind.SetBound True, "", "CLTCOD"
'    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
'    Call SetReDataCombo(oForm, sQry, oCombo)
'    oForm.Items("CLTCOD").DisplayDesc = True
'
'    oCombo.Select CLTCOD, psk_ByValue
    
    '사업장
    Call oForm.DataSources.UserDataSources.Add("CLTCOD", dt_SHORT_TEXT, 10)
    Call oForm.Items("CLTCOD").Specific.DataBind.SetBound(True, "", "CLTCOD")
    
    '시작일자
    Call oForm.DataSources.UserDataSources.Add("DocDateF", dt_DATE)
    Call oForm.Items("DocDateF").Specific.DataBind.SetBound(True, "", "DocDateF")
    
    '종료일자
    Call oForm.DataSources.UserDataSources.Add("DocDateT", dt_DATE)
    Call oForm.Items("DocDateT").Specific.DataBind.SetBound(True, "", "DocDateT")
    
    '부서
    Call oForm.DataSources.UserDataSources.Add("TeamCode", dt_SHORT_TEXT, 20)
    Call oForm.Items("TeamCode").Specific.DataBind.SetBound(True, "", "TeamCode")
    
    '사원번호
    Call oForm.DataSources.UserDataSources.Add("MSTCOD", dt_SHORT_TEXT, 20)
    Call oForm.Items("MSTCOD").Specific.DataBind.SetBound(True, "", "MSTCOD")
    
    '사원성명
    Call oForm.DataSources.UserDataSources.Add("MSTNAM", dt_SHORT_TEXT, 50)
    Call oForm.Items("MSTNAM").Specific.DataBind.SetBound(True, "", "MSTNAM")
    
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
    
PH_PY018_CreateItems_Error:
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY018_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY018_EnableMenus()
    
    On Error GoTo PH_PY018_EnableMenus_Error

    Call oForm.EnableMenu("1283", False)     '//제거
    Call oForm.EnableMenu("1284", False)    '//취소
    Call oForm.EnableMenu("1293", False)     '//행삭제
    
    Exit Sub
    
PH_PY018_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY018_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY018_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY018_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY018_FormItemEnabled
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY018_FormItemEnabled
        oForm.Items("Code").Specific.VALUE = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY018_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY018_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY018_FormItemEnabled()
    Dim oCombo      As SAPbouiCOM.ComboBox
    
    On Error GoTo PH_PY018_FormItemEnabled_Error
    
    
    Call oForm.Freeze(True)
    If (oForm.Mode = fm_ADD_MODE) Then

        Call CLTCOD_Select(oForm, "CLTCOD")

        Call oForm.EnableMenu("1281", False)     '//문서찾기
        Call oForm.EnableMenu("1282", True)    '//문서추가
        
    ElseIf (oForm.Mode = fm_FIND_MODE) Then
        
        Call CLTCOD_Select(oForm, "CLTCOD")

        Call oForm.EnableMenu("1281", False)    '//문서찾기
        Call oForm.EnableMenu("1282", False)     '//문서추가
        
    ElseIf (oForm.Mode = fm_OK_MODE) Then
        
        Call CLTCOD_Select(oForm, "CLTCOD")
        
        Call oForm.EnableMenu("1281", False)     '//문서찾기
        Call oForm.EnableMenu("1282", False)     '//문서추가

    End If
    Call oForm.Freeze(False)
    Exit Sub
    
PH_PY018_FormItemEnabled_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY018_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
                    If PH_PY018_DataValidCheck = True Then
                        Call PH_PY018_DataFind
                    Else
                        BubbleEvent = False
                    End If
                End If
                If pval.ItemUID = "Btn_Save" Then
                    If PH_PY018_DataSave = False Then
                        BubbleEvent = False
                    End If
                End If
                If pval.ItemUID = "Btn01" Then
                    If oForm.DataSources.DataTables.Item(0).Rows.Count > 0 Then
                        oForm.Freeze True
                        For i = 0 To oForm.DataSources.DataTables.Item(0).Rows.Count - 1
                            If oGrid1.DataTable.GetValue(4, i) = "Y" Then
                                oGrid1.DataTable.Columns(4).Cells.Item(i).VALUE = "N"
                            Else
                                oGrid1.DataTable.Columns(4).Cells.Item(i).VALUE = "Y"
                            End If
                        Next
                        oForm.Freeze False
                    End If
                End If
            ElseIf pval.BeforeAction = False Then
                If oForm.Mode = fm_ADD_MODE Then
                    If pval.ActionSuccess = True Then
'                        Call PH_PY018_FormItemEnabled
                    End If
                ElseIf oForm.Mode = fm_UPDATE_MODE Then
                    If pval.ActionSuccess = True Then
'                        Call PH_PY018_FormItemEnabled
                    End If
                ElseIf oForm.Mode = fm_OK_MODE Then
                    If pval.ActionSuccess = True Then
'                        Call PH_PY018_FormItemEnabled
                    End If
                End If

            End If
     
'----------------------------------------------------------
        Case et_KEY_DOWN: '//2
            
            If pval.BeforeAction = True Then
                If pval.CharPressed = 9 Then
                    If pval.ItemUID = "MSTCOD" Then
                        If oForm.Items("MSTCOD").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    End If
                End If
            End If
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
                    If pval.ItemUID = "CLTCOD" Then
                        
                        Dim CLTCOD As String

                        CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE)

                        If oForm.Items("TeamCode").Specific.ValidValues.Count > 0 Then
                            For i = oForm.Items("TeamCode").Specific.ValidValues.Count - 1 To 0 Step -1
                                Call oForm.Items("TeamCode").Specific.ValidValues.Remove(i, psk_Index)
                            Next i
                        End If

                        '부서콤보세팅
                        Call oForm.Items("TeamCode").Specific.ValidValues.Add("%", "전체")
                        sQry = "            SELECT      U_Code AS [Code],"
                        sQry = sQry & "                 U_CodeNm As [Name]"
                        sQry = sQry & "  FROM       [@PS_HR200L]"
                        sQry = sQry & "  WHERE      Code = '1'"
                        sQry = sQry & "                 AND U_UseYN = 'Y'"
                        sQry = sQry & "                 AND U_Char2 = '" & CLTCOD & "'"
                        sQry = sQry & "  ORDER BY  U_Seq"
                        Call MDC_SetMod.Set_ComboList(oForm.Items("TeamCode").Specific, sQry, "", False, False)
                        Call oForm.Items("TeamCode").Specific.Select(0, psk_Index)
                        
                        
                    End If
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
                    Select Case pval.ItemUID
                        Case "MSTCOD"
                            sQry = "SELECT U_FullName from [@PH_PY001A] Where Code = '" & oForm.Items("MSTCOD").Specific.VALUE & "'"
                            oRecordSet.DoQuery sQry
                            If oRecordSet.RecordCount > 0 Then
                                oForm.Items("MSTNAM").Specific.VALUE = oRecordSet.Fields(0).VALUE
                                
                            End If
'                        Case "Grid01"
'                            Select Case oForm.Items("CLTCOD").Specific.VALUE
'                                Case "1"
'                                    If oDS_PH_PY018.Columns.Item("DangerNu").Cells(pval.Row).VALUE <> 0 And oDS_PH_PY018.Columns.Item("DangerNu").Cells(pval.Row).VALUE <> 1 Then
'                                        oDS_PH_PY018.Columns.Item("DangerNu").Cells(pval.Row).VALUE = 0
'                                        Sbo_Application.SetStatusBarMessage "0 또는 1만 입력 가능합니다.", bmt_Short, True
'                                    End If
'                                    oDS_PH_PY018.Columns.Item("DangerCD").Cells(pval.Row).VALUE = ""
'                                Case "2"
'                                    If oDS_PH_PY018.Columns.Item("DangerNu").Cells(pval.Row).VALUE <> 0.5 And oDS_PH_PY018.Columns.Item("DangerNu").Cells(pval.Row).VALUE <> 1 Then
'                                        oDS_PH_PY018.Columns.Item("DangerNu").Cells(pval.Row).VALUE = 0
'                                        oDS_PH_PY018.Columns.Item("DangerCD").Cells(pval.Row).VALUE = ""
'                                        Sbo_Application.SetStatusBarMessage "0.5 또는 1 만 입력 가능합니다.", bmt_Short, True
'                                    Else
'                                        If oDS_PH_PY018.Columns.Item("DangerNu").Cells(pval.Row).VALUE >= 0.5 Then
'                                            oDS_PH_PY018.Columns.Item("DangerCD").Cells(pval.Row).VALUE = "56" '// '// 위해등급 6급
'                                        Else
'                                            oDS_PH_PY018.Columns.Item("DangerCD").Cells(pval.Row).VALUE = "" '//"56" '// 위해등급 6급
'                                        End If
'                                    End If
'                            End Select
                    End Select
                    
                End If
            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_MATRIX_LOAD: '//11
            If pval.BeforeAction = True Then
            ElseIf pval.BeforeAction = False Then
                Call PH_PY018_FormItemEnabled
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
                Set oDS_PH_PY018 = Nothing
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
                oForm.Items("MSTCOD").Specific.VALUE = ""
                oForm.Items("FullName").Specific.VALUE = ""
                oGrid1.DataTable.Clear
                
            Case "1288", "1289", "1290", "1291":
        End Select
    ElseIf (pval.BeforeAction = False) Then
        Select Case pval.MenuUID
            Case "1283":
                oForm.Mode = fm_ADD_MODE
                Call PH_PY018_FormItemEnabled

            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//문서찾기
                Call PH_PY018_FormItemEnabled
            Case "1282":    '//문서추가
                Call PH_PY018_FormItemEnabled

            Case "1288", "1289", "1290", "1291":
                Call PH_PY018_FormItemEnabled
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


Sub PH_PY018_FormClear()
On Error GoTo PH_PY018_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY018'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PH_PY018_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY018_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY018_DataValidCheck() As Boolean
On Error GoTo PH_PY018_DataValidCheck_Error
    PH_PY018_DataValidCheck = False
    Dim i       As Long
    Dim sQry    As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    If oForm.Items("CLTCOD").Specific.VALUE = "" Then
        Sbo_Application.SetStatusBarMessage "사업장은 필수입니다.", bmt_Short, True
        oForm.Items("CLTCOD").CLICK ct_Regular
        PH_PY018_DataValidCheck = False
        Exit Function
    End If
    
    PH_PY018_DataValidCheck = True
    
        
    Exit Function
    
    Set oRecordSet = Nothing
    
    
PH_PY018_DataValidCheck_Error:
    Set oRecordSet = Nothing
    PH_PY018_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY018_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Function PH_PY018_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY018_Validate_Error
    PH_PY018_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY018A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", bmt_Short, True
        PH_PY018_Validate = False
        GoTo PH_PY018_Validate_Exit
    End If
'
    If ValidateType = "수정" Then

    ElseIf ValidateType = "행삭제" Then

    ElseIf ValidateType = "취소" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY018_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY018_Validate_Error:
    PH_PY018_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY018_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Private Sub PH_PY018_DataFind()
    Dim i           As Long
    Dim iRow        As Long
    Dim sQry        As String
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY018_DataFind_Error
    
    Dim CLTCOD As String
    Dim DocDateF As String
    Dim DocDateT As String
    Dim TeamCode As String
    Dim MSTCOD As String
    
    CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE)
    DocDateF = oForm.Items("DocDateF").Specific.VALUE
    DocDateT = oForm.Items("DocDateT").Specific.VALUE
    TeamCode = oForm.Items("TeamCode").Specific.VALUE
    MSTCOD = oForm.Items("MSTCOD").Specific.VALUE
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    sQry = "Exec PH_PY018 '" & CLTCOD & "','" & DocDateF & "','" & DocDateT & "','" & TeamCode & "','" & MSTCOD & "'"

    Call oDS_PH_PY018.ExecuteQuery(sQry)
    
    iRow = oForm.DataSources.DataTables.Item(0).Rows.Count
    
    Call PH_PY018_TitleSetting(iRow)
    
    Set oRecordSet = Nothing
    Exit Sub
    
PH_PY018_DataFind_Error:
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "PH_PY018_DataFind_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PH_PY018_DataSave() As Boolean
    Dim i           As Long
    Dim ShiftDat    As String
    Dim sQry        As String
    Dim CLTCOD      As String
    
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    On Error GoTo PH_PY018_DataSave_Error
    
    PH_PY018_DataSave = False

    CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE)
    If oForm.DataSources.DataTables.Item(0).Rows.Count > 0 Then
        For i = 0 To oForm.DataSources.DataTables.Item(0).Rows.Count - 1
'            oDS_PH_PY018.Columns.Item("Code").Cells(i).Value
            
            sQry = " UPDATE ZPH_PY008 SET Attend = '" & oDS_PH_PY018.Columns.Item("Attend").Cells(i).VALUE & "'"
            sQry = sQry & " WHERE CLTCOD = '" & CLTCOD & "'"
            sQry = sQry & " And PosDate = '" & oDS_PH_PY018.Columns.Item("PosDate").Cells(i).VALUE & "'"
            sQry = sQry & " And MSTCOD = '" & oDS_PH_PY018.Columns.Item("MSTCOD").Cells(i).VALUE & "'"
            oRecordSet.DoQuery sQry
            
            
            
        Next i
        Sbo_Application.SetStatusBarMessage "연봉제휴일근무 수정변경되었습니다.", bmt_Short, False
        PH_PY018_DataSave = True
    Else
        Sbo_Application.SetStatusBarMessage "데이터가 존재하지 않습니다.", bmt_Short, True
    End If
    
    Set oRecordSet = Nothing
    Exit Function
    
PH_PY018_DataSave_Error:
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "PH_PY018_DataSave_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY018_TitleSetting(iRow As Long)
    Dim i               As Long
    Dim j               As Long
    Dim sQry            As String
     
    Dim COLNAM(4)       As String
    Dim CLTCOD          As String
    
    Dim oColumn         As SAPbouiCOM.EditTextColumn
    Dim oComboCol       As SAPbouiCOM.ComboBoxColumn
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
     
    On Error GoTo Error_Message
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    oForm.Freeze True
    

    COLNAM(0) = "일자"
    COLNAM(1) = "사번"
    COLNAM(2) = "성명"
    COLNAM(3) = "요일"
    COLNAM(4) = "휴일근무"
    
    CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE)
    
    
    For i = 0 To UBound(COLNAM)
        oGrid1.Columns.Item(i).TitleObject.Caption = COLNAM(i)
        
        Select Case COLNAM(i)
            Case "휴일근무"
                oGrid1.Columns.Item(i).Editable = True
                oGrid1.Columns.Item(i).Type = gct_ComboBox
                Set oComboCol = oGrid1.Columns.Item("Attend")
                
                oComboCol.ValidValues.Add "Y", "근무Y"
                oComboCol.ValidValues.Add "N", "N"
    
                oComboCol.DisplayType = cdt_Description
            Case Else
                oGrid1.Columns.Item(i).Editable = False
        End Select
        
    Next i
    
    oGrid1.AutoResizeColumns

    oForm.Freeze False

    Set oColumn = Nothing

    Exit Sub

Error_Message:
    oForm.Freeze False
    Set oColumn = Nothing
    Sbo_Application.SetStatusBarMessage "PH_PY018_TitleSetting Error : " & Space(10) & Err.Description, bmt_Short, True
End Sub




