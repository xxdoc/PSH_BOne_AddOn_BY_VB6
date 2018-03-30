VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY123"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY123.cls
'//  Module         : 급여관리 > 급여계산
'//  Desc           : 가압류등록
'//********************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form

Public oGrid1           As SAPbouiCOM.Grid
Public oDS_PH_PY123     As SAPbouiCOM.DataTable


Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long
Private oWorkType       As String


Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY123.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY123_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY123"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
'    oForm.DataBrowser.BrowseBy = "Code"
        
    oForm.PaneLevel = 1
    oForm.Freeze True
    Call PH_PY123_CreateItems
    Call PH_PY123_FormItemEnabled
    Call PH_PY123_EnableMenus
'    Call PH_PY123_SetDocument(oFromDocEntry01)
'    Call PH_PY123_FormResize
    
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

Private Function PH_PY123_CreateItems() As Boolean

    Dim sQry        As String
    Dim i           As Long
    Dim CLTCOD      As String
    
    
    Dim oEdit       As SAPbouiCOM.EditText
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns    As SAPbouiCOM.Columns
    Dim optBtn      As SAPbouiCOM.OptionBtn
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY123_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Set oGrid1 = oForm.Items("Grid01").Specific
    
    oForm.DataSources.DataTables.Add ("PH_PY123")
    
    oGrid1.DataTable = oForm.DataSources.DataTables.Item("PH_PY123")
    Set oDS_PH_PY123 = oForm.DataSources.DataTables.Item("PH_PY123")
    oGrid1.SelectionMode = ms_Auto

    '//----------------------------------------------------------------------------------------------
    '// 기본사항
    '//----------------------------------------------------------------------------------------------
    
    
    oForm.Items("CLTCOD").DisplayDesc = True
    
    
    Call oForm.DataSources.UserDataSources.Add("SMSTCOD", dt_SHORT_TEXT, 10)
    oForm.Items("SMSTCOD").Specific.DataBind.SetBound True, "", "SMSTCOD"
    
    Call oForm.DataSources.UserDataSources.Add("SFullName", dt_SHORT_TEXT, 10)
    oForm.Items("SFullName").Specific.DataBind.SetBound True, "", "SFullName"
    
    Call oForm.DataSources.UserDataSources.Add("MSTCOD", dt_SHORT_TEXT, 10)
    oForm.Items("MSTCOD").Specific.DataBind.SetBound True, "", "MSTCOD"
    
    Call oForm.DataSources.UserDataSources.Add("FullName", dt_SHORT_TEXT, 10)
    oForm.Items("FullName").Specific.DataBind.SetBound True, "", "FullName"
    '//일자
    Call oForm.DataSources.UserDataSources.Add("DocDate", dt_DATE, 8)
    oForm.Items("DocDate").Specific.DataBind.SetBound True, "", "DocDate"
    
    '//가압류금액
    Call oForm.DataSources.UserDataSources.Add("GaapAmt", dt_SUM, 10)
    oForm.Items("GaapAmt").Specific.DataBind.SetBound True, "", "GaapAmt"
    
    '//변제금액
    Call oForm.DataSources.UserDataSources.Add("BenjAmt", dt_SUM, 10)
    oForm.Items("BenjAmt").Specific.DataBind.SetBound True, "", "BenjAmt"
    
    
    Call oForm.DataSources.UserDataSources.Add("Comments", dt_SHORT_TEXT, 30)
    oForm.Items("Comments").Specific.DataBind.SetBound True, "", "Comments"
    
    
'    '// 구분
    Set oCombo = oForm.Items("Gubun").Specific
    oCombo.ValidValues.Add "1", "가압류등록"
    oCombo.ValidValues.Add "2", "변제등록"
    oCombo.ValidValues.Add "3", "종결처리"
    oForm.Items("Gubun").DisplayDesc = True
    
    '//급여구분
    Set oCombo = oForm.Items("JOBTYP").Specific
    oCombo.ValidValues.Add "", ""
    oCombo.ValidValues.Add "1", "급여"
    oCombo.ValidValues.Add "2", "상여"
    oForm.Items("JOBTYP").DisplayDesc = True
    
    
    oForm.Update
    
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
    
PH_PY123_CreateItems_Error:
    
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY123_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY123_EnableMenus()
    
    On Error GoTo PH_PY123_EnableMenus_Error

    Call oForm.EnableMenu("1283", False)     '//제거
    Call oForm.EnableMenu("1284", False)    '//취소
    Call oForm.EnableMenu("1293", False)     '//행삭제
    
    Exit Sub
    
PH_PY123_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY123_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY123_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY123_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY123_FormItemEnabled
        Call PH_PY123_AddMatrixRow
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY123_FormItemEnabled
        oForm.Items("Code").Specific.VALUE = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY123_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY123_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY123_FormItemEnabled()
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim sQry        As String
    Dim i           As Long
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    
    Dim CLTCOD As String
    Dim sPosDate As String
    
    On Error GoTo PH_PY123_FormItemEnabled_Error
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Call oForm.Freeze(True)
    If (oForm.Mode = fm_ADD_MODE) Then
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        
        Set oCombo = oForm.Items("CLTCOD").Specific
        CLTCOD = MDC_SetMod.Get_ReData("Branch", "USER_CODE", "OUSR", "'" & oCompany.UserName & "'")
        oCombo.Select CLTCOD, psk_ByValue
        
        oForm.Items.Item("MSTCOD").Enabled = True
        oForm.Items.Item("FullName").Enabled = False
        oForm.Items.Item("CLTCOD").Enabled = True
        
        oForm.Items("DocDate").Specific.VALUE = Format(Now, "yyyymmdd")
        
        oForm.Items("SMSTCOD").Specific.VALUE = ""
        oForm.Items("SFullName").Specific.VALUE = ""
        
        oForm.Items("MSTCOD").Specific.VALUE = ""
        oForm.Items("FullName").Specific.VALUE = ""
        
        '//구분
        Set oCombo = oForm.Items("Gubun").Specific
        oCombo.Select "1", psk_ByValue
        
        '//급여구분
        Set oCombo = oForm.Items("JOBTYP").Specific
        oCombo.Select "", psk_ByValue
        
        oForm.DataSources.UserDataSources.Item("Comments").VALUE = ""
        
        
        oForm.DataSources.UserDataSources.Item("GaapAmt").VALUE = 0
        oForm.DataSources.UserDataSources.Item("BenjAmt").VALUE = 0
        
        
        
        Call oForm.EnableMenu("1281", False)     '//문서찾기
        Call oForm.EnableMenu("1282", True)    '//문서추가

    ElseIf (oForm.Mode = fm_FIND_MODE) Then
        
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        
        Call oForm.EnableMenu("1281", False)    '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가
        
        
    ElseIf (oForm.Mode = fm_OK_MODE) Then
        oForm.Items.Item("CLTCOD").Enabled = False
        
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD", False)
        
        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가
        
    End If
    Set oCombo = Nothing
    Call oForm.Freeze(False)
    Exit Sub
    
PH_PY123_FormItemEnabled_Error:
    Set oCombo = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY123_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Public Sub Raise_FormItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
    Dim sQry        As String
    Dim i           As Long
    Dim Chkcnt      As Long
    Dim tSex        As String
    Dim tBrith      As String
    Dim Day         As String
    Dim ActCode     As String
    Dim MSTCOD      As String
    Dim CLTCOD      As String
    Dim YY          As String
    Dim ymd         As String
    Dim JanQty      As Double
    
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
                    If PH_PY123_DataValidCheck = False Then
                        BubbleEvent = False
                    End If
                End If
                
                If pval.ItemUID = "Btn_ret" Then

                    Call PH_PY123_MTX01
                End If
                
                
                
                If pval.ItemUID = "Btn01" Then
                    Call PH_PY123_SAVE
                End If
                
                
                If pval.ItemUID = "Btn_del" Then
                    Call PH_PY123_Delete
                End If
                
                
            ElseIf pval.BeforeAction = False Then
                Select Case pval.ItemUID
                Case "1"
                    If oForm.Mode = fm_ADD_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY123_FormItemEnabled
                            Call PH_PY123_AddMatrixRow
                        End If
                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY123_FormItemEnabled
                            Call PH_PY123_AddMatrixRow
                        End If
                    ElseIf oForm.Mode = fm_OK_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY123_FormItemEnabled
                        End If
                    End If
                
                End Select
            End If
     
'----------------------------------------------------------
        Case et_KEY_DOWN: '//2
            If pval.BeforeAction = True Then
                If pval.CharPressed = 9 Then
                    If pval.ItemUID = "SMSTCOD" Then
                        If oForm.Items("SMSTCOD").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    ElseIf pval.ItemUID = "MSTCOD" Then
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
                    Select Case pval.ItemUID
                        
                        Case "Gubun"
                            '//구분
                            '변제등록일때
                            If oForm.Items("Gubun").Specific.VALUE = "3" Then
                            End If
                            
                        End Select
                    
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
                            PH_PY123_MTX02 pval.ItemUID, pval.Row, pval.ColUID
                            
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
                Select Case pval.ItemUID
                Case "Grid01"
                    If pval.Row >= 0 Then
                        Select Case pval.ItemUID
                        Case "Grid01"
                            'Call oMat1.SelectRow(pval.Row, True, False)
                            
                        End Select
                        
                    End If
                End Select
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

                        Case "SMSTCOD"
                            Call oForm.Freeze(True)
                            MSTCOD = oForm.Items("SMSTCOD").Specific.VALUE
                            sQry = "SELECT U_FullName from [@PH_PY001A] Where Code = '" & MSTCOD & "'"
                            oRecordSet.DoQuery sQry
                            If oRecordSet.RecordCount > 0 Then
                                oForm.Items("SFullName").Specific.VALUE = oRecordSet.Fields(0).VALUE
                            Else
                                oForm.Items("SFullName").Specific.VALUE = ""
                            End If
                            Call oForm.Freeze(False)
                        Case "MSTCOD"
                            Call oForm.Freeze(True)
                            MSTCOD = oForm.Items("MSTCOD").Specific.VALUE
                            sQry = "SELECT U_FullName, U_TeamCode, U_RspCode, U_ShiftDat, U_GNMUJO from [@PH_PY001A] Where Code = '" & MSTCOD & "'"
                            oRecordSet.DoQuery sQry
                            If oRecordSet.RecordCount > 0 Then
                                oForm.Items("FullName").Specific.VALUE = oRecordSet.Fields(0).VALUE
                                
                            Else
                                oForm.Items("FullName").Specific.VALUE = ""
                            End If
                            Call oForm.Freeze(False)
                    End Select
                 
                End If
            End If
'            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_MATRIX_LOAD: '//11
            If pval.BeforeAction = True Then
            ElseIf pval.BeforeAction = False Then
'                oMat1.LoadFromDataSource
'                Call PH_PY123_AddMatrixRow
                
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
                Set oDS_PH_PY123 = Nothing
                
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
                oForm.Items.Item("79").Width = oForm.Items.Item("KUKGRD").Left + oForm.Items.Item("KUKGRD").Width - oForm.Items.Item("79").Left + 10
                oForm.Items.Item("79").Height = oForm.Items.Item("80").Height
                
                oForm.Items.Item("77").Width = oForm.Items.Item("BUYN20").Left + oForm.Items.Item("BUYN20").Width - oForm.Items.Item("77").Left + 16
                oForm.Items.Item("77").Height = oForm.Items.Item("78").Height
            
                
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
'                    Call MDC_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY123A", "Code")
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
'                Call PH_PY123_FormItemEnabled
        End Select
    ElseIf (pval.BeforeAction = False) Then
        Select Case pval.MenuUID
            Case "1283":
                oForm.Mode = fm_ADD_MODE
                Call PH_PY123_FormItemEnabled
'                Call PH_PY123_AddMatrixRow
            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//문서찾기
                Call PH_PY123_FormItemEnabled
'                Call PH_PY123_AddMatrixRow
                oForm.Items("Code").CLICK ct_Regular
            Case "1282":    '//문서추가
                Call PH_PY123_FormItemEnabled
'                Call PH_PY123_AddMatrixRow
            Case "1288", "1289", "1290", "1291":
                Call PH_PY123_FormItemEnabled
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

Sub PH_PY123_AddMatrixRow()
    Dim oRow    As Long
    
    On Error GoTo PH_PY123_AddMatrixRow_Error
    
    Call oForm.Freeze(True)
    
'    '//[Mat1]
'    oMat1.FlushToDataSource
'    oRow = oMat1.VisualRowCount
'
'    If oMat1.VisualRowCount > 0 Then
'        If Trim(oDS_PH_PY123A.GetValue("U_MSTCOD", oRow - 1)) <> "" Then
'            If oDS_PH_PY123A.Size <= oMat1.VisualRowCount Then
'                oDS_PH_PY123A.InsertRecord (oRow)
'            End If
'            oDS_PH_PY123A.Offset = oRow
'            oDS_PH_PY123A.setValue "U_LineNum", oRow, oRow + 1
'            oMat1.LoadFromDataSource
'        Else
'            oDS_PH_PY123A.Offset = oRow - 1
'            oDS_PH_PY123A.setValue "U_LineNum", oRow - 1, oRow
'            oMat1.LoadFromDataSource
'        End If
'    ElseIf oMat1.VisualRowCount = 0 Then
'        oDS_PH_PY123A.Offset = oRow
'        oDS_PH_PY123A.setValue "U_LineNum", oRow, oRow + 1
'        oMat1.LoadFromDataSource
'    End If

    

    Call oForm.Freeze(False)
    Exit Sub
PH_PY123_AddMatrixRow_Error:

    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY123_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY123_FormClear()
On Error GoTo PH_PY123_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY123'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PH_PY123_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY123_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY123_DataValidCheck() As Boolean
On Error GoTo PH_PY123_DataValidCheck_Error
    PH_PY123_DataValidCheck = False
    Dim i       As Long
    Dim j       As Long
    
    Dim sQry    As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
'    '//----------------------------------------------------------------------------------
'    '//기본사항 탭
'    '//----------------------------------------------------------------------------------
'    If Trim(oDS_PH_PY123A.GetValue("Code", 0)) = "" Then
'        Sbo_Application.SetStatusBarMessage "사원번호는 필수입니다.", bmt_Short, True
'        oForm.Items("Code").CLICK ct_Regular
'        PH_PY123_DataValidCheck = False
'        Exit Function
'    End If
'
'    If oForm.Mode = fm_ADD_MODE Then
'        '//사원번호 중복 체크 (PH_PY123A)
'        sQry = "SELECT Code FROM [@PH_PY123A] WHERE Code = '" & Trim(oDS_PH_PY123A.GetValue("Code", 0)) & "'"
'        oRecordSet.DoQuery sQry
'
'        If oRecordSet.RecordCount > 0 Then
'            Sbo_Application.SetStatusBarMessage "[PH_PY123A] 이미 존재하는 사원번호입니다", bmt_Short, True
'            oForm.Items("Code").CLICK ct_Regular
'            PH_PY123_DataValidCheck = False
'            Exit Function
'        End If
'
'        '//사원번호 중복 체크(OHEM)
'        sQry = "SELECT U_MSTCOD FROM OHEM WHERE U_MSTCOD = '" & Trim(oDS_PH_PY123A.GetValue("Code", 0)) & "'"
'        oRecordSet.DoQuery sQry
'
'        If oRecordSet.RecordCount > 0 Then
'            Sbo_Application.SetStatusBarMessage "[OHEM] 이미 존재하는 사원번호입니다", bmt_Short, True
'            oForm.Items("Code").CLICK ct_Regular
'            PH_PY123_DataValidCheck = False
'            Exit Function
'        End If
'    End If
'    '// Code = Name 동일
'    Call oDS_PH_PY123A.setValue("Name", 0, oDS_PH_PY123A.GetValue("Code", 0))
'
'
'
'    oMat1.FlushToDataSource
'
'    '// Matrix 마지막 행 삭제(DB 저장시)
'    If oDS_PH_PY123A.Size > 1 Then oDS_PH_PY123B.RemoveRecord (oDS_PH_PY123B.Size - 1)
'
'    oMat1.LoadFromDataSource
'
'    PH_PY123_DataValidCheck = True
    
    Exit Function
    
    Set oRecordSet = Nothing
    
    
PH_PY123_DataValidCheck_Error:
    Set oRecordSet = Nothing
    PH_PY123_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY123_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY123_MTX01()

    '//메트릭스에 데이터 로드
    
    Dim i       As Long
    Dim Chkcnt  As Long
    Dim sQry    As String
    Dim iRow        As Long
    
    Dim Param01 As String
    Dim Param02 As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY123_MTX01_Error

    Call oForm.Freeze(True)
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    Param01 = Trim(oForm.Items("CLTCOD").Specific.VALUE)
    Param02 = oForm.Items("SMSTCOD").Specific.VALUE
    
    
    sQry = "EXEC PH_PY123_01 '" & Param01 & "', '" & Param02 & "'"
    Call oDS_PH_PY123.ExecuteQuery(sQry)
    
    iRow = oForm.DataSources.DataTables.Item(0).Rows.Count
'
    Call PH_PY123_TitleSetting(iRow)
    
    
    oForm.Update
    
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Sub
PH_PY123_MTX01_Exit:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Sub
PH_PY123_MTX01_Error:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY123_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub
Private Sub PH_PY123_MTX02(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)


    '//메트릭스 자료를 head에 로드
    
    Dim i       As Long
    Dim sQry    As String
    Dim sRow As Long
    
    Dim Param01 As String
    Dim Param02 As String
    Dim Param03 As String
    Dim Param04 As String
    
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY123_MTX02_Error

    Call oForm.Freeze(True)
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    sRow = oRow
    
    
    Param01 = Trim(oForm.Items("CLTCOD").Specific.VALUE)
    Param02 = oDS_PH_PY123.Columns.Item("DocDate").Cells(oRow).VALUE
    Param03 = oDS_PH_PY123.Columns.Item("MSTCOD").Cells(oRow).VALUE
    Param04 = oDS_PH_PY123.Columns.Item("Gubun").Cells(oRow).VALUE
    
    If Trim(Param02) = "" Then
    
        Set oCombo = oForm.Items("Gubun").Specific
        oCombo.Select "", psk_ByValue
        
        Set oCombo = oForm.Items("JOBTYP").Specific
        oCombo.Select "", psk_ByValue
        
        oForm.Items("MSTCOD").Specific.VALUE = ""
        oForm.Items("FullName").Specific.VALUE = ""
        
        oForm.DataSources.UserDataSources.Item("DocDate").VALUE = ""
        oForm.DataSources.UserDataSources.Item("Comments").VALUE = ""
        
        oForm.DataSources.UserDataSources.Item("GaapAmt").VALUE = 0
        oForm.DataSources.UserDataSources.Item("BenjAmt").VALUE = 0
    
        
        GoTo PH_PY123_MTX02_Exit
    End If
    sQry = "EXEC PH_PY123_02 '" & Param01 & "', '" & Param02 & "', '" & Param03 & "', '" & Param04 & "'"
    Call oRecordSet.DoQuery(sQry)
    
    If (oRecordSet.RecordCount = 0) Then
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
        GoTo PH_PY123_MTX02_Exit
    End If
    
    oForm.DataSources.UserDataSources.Item("DocDate").VALUE = oRecordSet.Fields("DocDate").VALUE
        
    Set oCombo = oForm.Items("Gubun").Specific
    oCombo.Select oRecordSet.Fields("Gubun").VALUE, psk_ByValue
        
    Set oCombo = oForm.Items("JOBTYP").Specific
    oCombo.Select oRecordSet.Fields("JOBTYP").VALUE, psk_ByValue
        
        
    oForm.DataSources.UserDataSources.Item("MSTCOD").VALUE = oRecordSet.Fields("MSTCOD").VALUE
    oForm.DataSources.UserDataSources.Item("FullName").VALUE = oRecordSet.Fields("FullName").VALUE
    oForm.DataSources.UserDataSources.Item("GaapAmt").VALUE = oRecordSet.Fields("GaapAmt").VALUE
    oForm.DataSources.UserDataSources.Item("BenjAmt").VALUE = oRecordSet.Fields("BenjAmt").VALUE
    oForm.DataSources.UserDataSources.Item("Comments").VALUE = oRecordSet.Fields("Comments").VALUE
    
    
    oForm.Update
    
    
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Sub
PH_PY123_MTX02_Exit:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    
    Exit Sub
PH_PY123_MTX02_Error:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY123_MTX02_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY123_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY123_Validate_Error
    PH_PY123_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY123A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", bmt_Short, True
        PH_PY123_Validate = False
        GoTo PH_PY123_Validate_Exit
    End If
'
    If ValidateType = "수정" Then

    ElseIf ValidateType = "행삭제" Then

    ElseIf ValidateType = "취소" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY123_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY123_Validate_Error:
    PH_PY123_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY123_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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


Private Sub PH_PY123_SAVE()

    '//데이타 저장
    Dim oCombo      As SAPbouiCOM.ComboBox
    
    Dim i       As Long
    Dim sQry    As String
    Dim ErrNum As Integer
    
    Dim CLTCOD$, DocDate$, MSTCOD$, Gubun$, JOBTYP$, Comments$
    Dim GaapAmt As Double
    Dim BenjAmt As Double
    
    Dim Param01 As String
    Dim Param02 As String
    Dim Param03 As String
    Dim Param04 As String
    Dim Param05 As String
    Dim Param06 As String
    Dim Param07 As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY123_SAVE_Error

    Call oForm.Freeze(True)
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    
    MSTCOD = oForm.Items("MSTCOD").Specific.VALUE
    CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE)
    DocDate = oForm.Items("DocDate").Specific.VALUE
    Gubun = oForm.Items("Gubun").Specific.VALUE
    JOBTYP = oForm.Items("JOBTYP").Specific.VALUE
    GaapAmt = oForm.Items("GaapAmt").Specific.VALUE
    BenjAmt = oForm.Items("BenjAmt").Specific.VALUE
    Comments = oForm.Items("Comments").Specific.VALUE
            
    If Trim(JOBTYP) <> "" Then
        Call MDC_Com.MDC_GF_Message("급여계산 된 자료는 수정할 수 없습니다.", "E")
        GoTo PH_PY123_SAVE_Exit
    End If
                
    If Gubun = "" Then
        Call MDC_Com.MDC_GF_Message("구분은 필수 입력사항입니다.", "E")
        GoTo PH_PY123_SAVE_Exit
    End If
    
    sQry = "Select Cnt = Count(*) From ZPH_PY123 Where CLTCOD = '" + CLTCOD + "' AND MSTCOD = '" + MSTCOD + "'"
    sQry = sQry + " And DocDate = '" + DocDate + "' and Gubun = '" + Gubun + "'"
    oRecordSet.DoQuery sQry
    
    If oRecordSet.Fields(0).VALUE <= 0 Then
        '신규
        sQry = "INSERT INTO ZPH_PY123"
        sQry = sQry & " ("
        sQry = sQry & " CLTCOD,"
        sQry = sQry & " DocDate,"
        sQry = sQry & " MSTCOD,"
        sQry = sQry & " Gubun,"
        sQry = sQry & " JOBTYP,"
        sQry = sQry & " GaapAmt,"
        sQry = sQry & " BenjAmt,"
        sQry = sQry & " Comments"
        sQry = sQry & " ) "
        sQry = sQry & "VALUES("
        sQry = sQry & "'" & CLTCOD & "',"
        sQry = sQry & "'" & DocDate & "',"
        sQry = sQry & "'" & MSTCOD & "',"
        sQry = sQry & "'" & Gubun & "',"
        sQry = sQry & "'" & JOBTYP & "',"
        sQry = sQry & GaapAmt & ","
        sQry = sQry & BenjAmt & ","
        sQry = sQry & "'" & Comments & "'"
        sQry = sQry & ")"
        oRecordSet.DoQuery sQry
        
        oForm.Mode = fm_ADD_MODE
    Else
        '수정
        sQry = "Update ZPH_PY123"
        sQry = sQry & " Set DocDate = '" & DocDate & "',"
        sQry = sQry & " Gubun = '" & Gubun & "',"
        sQry = sQry & " MSTCOD = '" & MSTCOD & "',"
        sQry = sQry & " GaapAmt = " & GaapAmt & ","
        sQry = sQry & " BenjAmt = " & BenjAmt & ","
        sQry = sQry & " Comments = '" & Comments & "'"
        sQry = sQry & " Where CLTCOD = '" & CLTCOD & "'"
        sQry = sQry & " And DocDate = '" & DocDate & "'"
        sQry = sQry & " And MSTCOD = '" & MSTCOD & "'"
        sQry = sQry & " And Gubun = '" & Gubun & "'"
        
        oRecordSet.DoQuery sQry
        
        oForm.Mode = fm_OK_MODE
    End If
    
    
    
    
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    
                    
    Call PH_PY123_MTX01
    
    Call PH_PY123_FormItemEnabled

    Exit Sub
    
PH_PY123_SAVE_Exit:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Sub
PH_PY123_SAVE_Error:
    Call oForm.Freeze(False)
    
    Set oRecordSet = Nothing
    
    Sbo_Application.SetStatusBarMessage "PH_PY123_SAVE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    
    
    
End Sub




Private Sub PH_PY123_Delete()
    '//선택된 자료 삭제
    Dim DocDate As String
    Dim CLTCOD As String
    Dim MSTCOD As String
    Dim Gubun As String
    Dim JOBTYP As String
    
    Dim i As Integer
    Dim sQry    As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY123_Delete_Error

    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    
    
    Call oForm.Freeze(True)
    
    CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE)
    
    MSTCOD = oForm.Items("MSTCOD").Specific.VALUE
    Gubun = oForm.Items("Gubun").Specific.VALUE
    DocDate = oForm.Items("DocDate").Specific.VALUE
    JOBTYP = oForm.Items("JOBTYP").Specific.VALUE
    
    If Trim(JOBTYP) <> "" Then
        Call MDC_Com.MDC_GF_Message("급여계산 된 자료는 삭제할 수 없습니다.", "E")
        GoTo PH_PY123_Delete_Exit
    End If
    
    
    If Sbo_Application.MessageBox(" 선택한자료를 삭제하시겠습니까? ?", "2", "예", "아니오") = "1" Then
        If oDS_PH_PY123.Rows.Count > 0 Then
                    
            sQry = "Delete From ZPH_PY123 Where CLTCOD = '" & CLTCOD & "' AND  DocDate = '" & DocDate & "' And MSTCOD = '" & MSTCOD & "' And Gubun = '" & Gubun & "'"
            oRecordSet.DoQuery sQry
            
        End If
    Else
        GoTo PH_PY123_Delete_Exit
    End If
    
    oForm.Mode = fm_ADD_MODE
    
    Call PH_PY123_MTX01
    
    Call PH_PY123_FormItemEnabled
    
    Call oForm.Freeze(False)
    
    Set oRecordSet = Nothing
    
    Exit Sub
PH_PY123_Delete_Exit:
    Set oRecordSet = Nothing
    
    Call oForm.Freeze(False)
    Exit Sub
PH_PY123_Delete_Error:
    Set oRecordSet = Nothing
    
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY123_Delete_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub PH_PY123_TitleSetting(iRow As Long)
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
    
    COLNAM(0) = "사업장"
    COLNAM(1) = "사번"
    COLNAM(2) = "성명"
    COLNAM(3) = "일자"
    COLNAM(4) = "구분"
    COLNAM(5) = "압류금액"
    COLNAM(6) = "급여구분"
    COLNAM(7) = "변제금액"
    COLNAM(8) = "비고"

    For i = 0 To UBound(COLNAM)
        oGrid1.Columns.Item(i).TitleObject.Caption = COLNAM(i)
        oGrid1.Columns.Item(i).Editable = False
        

        Select Case COLNAM(i)
            
            Case "구분"
                oGrid1.Columns.Item(i).Type = gct_ComboBox
                Set oComboCol = oGrid1.Columns.Item("Gubun")
                
                oComboCol.ValidValues.Add "1", "가압류등록"
                oComboCol.ValidValues.Add "2", "변제등록"
                oComboCol.ValidValues.Add "3", "종결처리"
                
                oComboCol.DisplayType = cdt_Description
            Case "급여구분    "
                oGrid1.Columns.Item(i).Type = gct_ComboBox
                Set oComboCol = oGrid1.Columns.Item("JOBTYP")
                
                oComboCol.ValidValues.Add "1", "급여"
                oComboCol.ValidValues.Add "2", "상여"
                
                oComboCol.DisplayType = cdt_Description
            
            Case Else
                oGrid1.Columns.Item(i).RightJustified = True

'            Case Else
'                oGrid1.Columns.Item(i).Editable = False
        End Select
        
    Next i
    
    oGrid1.AutoResizeColumns

    oForm.Freeze False

    Set oColumn = Nothing

    Exit Sub

Error_Message:
    oForm.Freeze False
    Set oColumn = Nothing
    Sbo_Application.SetStatusBarMessage "PH_PY123_TitleSetting Error : " & Space(10) & Err.Description, bmt_Short, True
End Sub


'Private Sub PH_PY123_View()
'
'    '//메트릭스에 근태이상자 로드
'
'    Dim i       As Long
'    Dim sQry    As String
'    Dim iRow        As Long
'
'    Dim PosDate As String
'    Dim CLTCOD As String
'
'    Dim oRecordSet As SAPbobsCOM.Recordset
'
'    On Error GoTo PH_PY123_View_Error
'
'    Call oForm.Freeze(True)
'
'    oMat1.Clear
''    oMat1.FlushToDataSource
''    oMat1.LoadFromDataSource
'
'    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
'
'    PosDate = oForm.Items("PosDate").Specific.Value
'    CLTCOD = oForm.Items("SCLTCOD").Specific.Value
'
'    sQry = "select PosDate = Convert(Char(8),p.PosDate,112),"
'    sQry = sQry + " TeamName = (SELECT U_CodeNm FROM [@PS_HR200L] where Code = '1' and  U_Code = p.TeamCode ),"
'    sQry = sQry + " RspName = (SELECT U_CodeNm FROM [@PS_HR200L] where Code = '2' and  U_Code = p.RspCode ),"
'    sQry = sQry + " MSTCOD = p.MSTCOD, p.FullName"
'    sQry = sQry + " from ZPH_PY123 p"
'    sQry = sQry + " Where p.PosDate    = '" & PosDate & "'"
'    sQry = sQry + " And p.CLTCOD = '" & CLTCOD & "'"
'    sQry = sQry + " and p.Confirm  = 'Y'"
'    sQry = sQry + " and p.Attend = 'N'"
'
'    oRecordSet.DoQuery sQry
'    If oRecordSet.RecordCount = 0 Then
'        Call MDC_Com.MDC_GF_Message("근태이상자 자료가 없습니다..", "E")
'        GoTo PH_PY123_View_Exit
'    End If
'
'    For i = 1 To oRecordSet.RecordCount
'        oMat1.AddRow (1)
'        oMat1.Columns("LineNum").Cells(i).Specific.Value = i
'        oMat1.Columns("PosDate").Cells(i).Specific.Value = oRecordSet.Fields(0).Value
'        oMat1.Columns("TeamName").Cells(i).Specific.Value = oRecordSet.Fields(1).Value
'        oMat1.Columns("RspName").Cells(i).Specific.Value = oRecordSet.Fields(2).Value
'        oMat1.Columns("MSTCOD").Cells(i).Specific.Value = oRecordSet.Fields(3).Value
'        oMat1.Columns("FullName").Cells(i).Specific.Value = oRecordSet.Fields(4).Value
'
''        If i <> 0 Then
''            oDS_PH_PYMat.InsertRecord (i)
''        End If
''        oDS_PH_PYMat.Offset = i
''        Call oDS_PH_PYMat.setValue("U_LineNum", i, i + 1)
''        Call oDS_PH_PYMat.setValue("U_ColReg01", i, oRecordSet.Fields(0).Value)
''        Call oDS_PH_PYMat.setValue("U_ColReg02", i, oRecordSet.Fields(1).Value)
''        Call oDS_PH_PYMat.setValue("U_ColReg03", i, oRecordSet.Fields(2).Value)
''        Call oDS_PH_PYMat.setValue("U_ColReg04", i, oRecordSet.Fields(3).Value)
''        Call oDS_PH_PYMat.setValue("U_ColReg05", i, oRecordSet.Fields(4).Value)
'        oRecordSet.MoveNext
''
'
'    Next
'
''    oMat1.LoadFromDataSource
'    oMat1.AutoResizeColumns
'    oForm.Items("Mat01").Visible = True
'    oForm.Update
'
'    Set oRecordSet = Nothing
'    Call oForm.Freeze(False)
'    Exit Sub
'PH_PY123_View_Exit:
'    Set oRecordSet = Nothing
'    Call oForm.Freeze(False)
'    Exit Sub
'PH_PY123_View_Error:
'    Set oRecordSet = Nothing
'    Call oForm.Freeze(False)
'    Sbo_Application.SetStatusBarMessage "PH_PY123_View_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
'End Sub


