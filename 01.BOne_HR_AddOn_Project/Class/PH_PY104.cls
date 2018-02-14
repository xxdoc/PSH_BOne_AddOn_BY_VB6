VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY104"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY104.cls
'//  Module         : 급여관리 > 급여관리
'//  Desc           : 고정수당공제금액일괄등록
'//********************************************************************************
Option Explicit

Public oFormUniqueID    As String
Public oForm            As SAPbouiCOM.Form

Public oGrid1           As SAPbouiCOM.Grid
Public oGrid2           As SAPbouiCOM.Grid
Public oDS_PH_PY104_01  As SAPbouiCOM.DataTable
Public oDS_PH_PY104_02  As SAPbouiCOM.DataTable

Public tSeqAll     As Long  '//그리드1의 체크 순번

Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long

Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY104.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY104_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY104"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE

        
    oForm.Freeze True
    Call PH_PY104_CreateItems
    Call PH_PY104_EnableMenus
    Call PH_PY104_SetDocument(oFromDocEntry01)
    
'    Call PH_PY104_FormResize
    
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

Private Function PH_PY104_CreateItems() As Boolean

    Dim sQry        As String
    Dim i           As Long
    
    Dim oCheck      As SAPbouiCOM.CheckBox
    Dim oEdit       As SAPbouiCOM.EditText
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns    As SAPbouiCOM.Columns
    Dim optBtn      As SAPbouiCOM.OptionBtn
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY104_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Set oGrid1 = oForm.Items("Grid1").Specific
    
    oForm.DataSources.DataTables.Add ("PH_PY104_01")
    oForm.DataSources.DataTables.Item("PH_PY104_01").Columns.Add "구분", ft_AlphaNumeric
    oForm.DataSources.DataTables.Item("PH_PY104_01").Columns.Add "코드", ft_AlphaNumeric
    oForm.DataSources.DataTables.Item("PH_PY104_01").Columns.Add "이름", ft_AlphaNumeric
    oForm.DataSources.DataTables.Item("PH_PY104_01").Columns.Add "선택", ft_AlphaNumeric
    oForm.DataSources.DataTables.Item("PH_PY104_01").Columns.Add "순서", ft_Float
    
    oGrid1.DataTable = oForm.DataSources.DataTables.Item("PH_PY104_01")
    Set oDS_PH_PY104_01 = oForm.DataSources.DataTables.Item("PH_PY104_01")
    
    Set oGrid2 = oForm.Items("Grid2").Specific
    
    oForm.DataSources.DataTables.Add ("PH_PY104_02")
    oForm.DataSources.DataTables.Item("PH_PY104_02").Columns.Add "체크", ft_AlphaNumeric
    oForm.DataSources.DataTables.Item("PH_PY104_02").Columns.Add "사번", ft_AlphaNumeric
    oForm.DataSources.DataTables.Item("PH_PY104_02").Columns.Add "이름", ft_AlphaNumeric

    oGrid2.DataTable = oForm.DataSources.DataTables.Item("PH_PY104_02")
    Set oDS_PH_PY104_02 = oForm.DataSources.DataTables.Item("PH_PY104_02")
    
    
    '//----------------------------------------------------------------------------------------------
    '// 기본사항
    '//----------------------------------------------------------------------------------------------
    
    '//사업장
    Call oForm.DataSources.UserDataSources.Add("CLTCOD", dt_SHORT_TEXT, 10)
    Set oCombo = oForm.Items("CLTCOD").Specific
    oCombo.DataBind.SetBound True, "", "CLTCOD"
'    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
'    Call SetReDataCombo(oForm, sQry, oCombo)
    oForm.Items("CLTCOD").DisplayDesc = True
    
    '//부서
    Call oForm.DataSources.UserDataSources.Add("TeamCode", dt_SHORT_TEXT, 10)
    Set oCombo = oForm.Items("TeamCode").Specific
    oCombo.DataBind.SetBound True, "", "TeamCode"
    oForm.Items("TeamCode").DisplayDesc = True
'
    '// 담당
    Call oForm.DataSources.UserDataSources.Add("RspCode", dt_SHORT_TEXT, 10)
    Set oCombo = oForm.Items("RspCode").Specific
    oCombo.DataBind.SetBound True, "", "RspCode"
    oForm.Items("RspCode").DisplayDesc = True
    
    '// 급여형태
    Call oForm.DataSources.UserDataSources.Add("PAYTYP", dt_SHORT_TEXT, 10)
    Set oCombo = oForm.Items("PAYTYP").Specific
    oCombo.DataBind.SetBound True, "", "PAYTYP"
    
    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P132' AND U_UseYN= 'Y' "
    Call SetReDataCombo(oForm, sQry, oCombo)
    oCombo.ValidValues.Add "%", "전체"
    oCombo.Select "%", psk_ByValue
    oForm.Items("PAYTYP").DisplayDesc = True

    '// 직급형태From
    Set oCombo = oForm.Items("JIGCODF").Specific
    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P129' ORDER BY U_Code "
    Call SetReDataCombo(oForm, sQry, oCombo)
    oCombo.ValidValues.Add "%", "전체"
    oCombo.Select "%", psk_ByValue
    oForm.Items("JIGCODF").DisplayDesc = True
    
    '// 직급형태To
    Set oCombo = oForm.Items("JIGCODT").Specific
    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P129' ORDER BY U_Code "
    Call SetReDataCombo(oForm, sQry, oCombo)
    oCombo.ValidValues.Add "%", "전체"
    oCombo.Select "%", psk_ByValue
    oForm.Items("JIGCODT").DisplayDesc = True
    
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
    
PH_PY104_CreateItems_Error:
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY104_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY104_EnableMenus()
    
    On Error GoTo PH_PY104_EnableMenus_Error

    Call oForm.EnableMenu("1283", True)     '//제거
    Call oForm.EnableMenu("1284", False)    '//취소
    Call oForm.EnableMenu("1293", True)     '//행삭제
    
    Exit Sub
    
PH_PY104_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY104_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY104_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY104_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY104_FormItemEnabled
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY104_FormItemEnabled
        oForm.Items("Code").Specific.Value = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY104_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY104_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY104_FormItemEnabled()
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim i           As Long
    Dim sQry        As String
    On Error GoTo PH_PY104_FormItemEnabled_Error
    
    
    Call oForm.Freeze(True)
    If (oForm.Mode = fm_ADD_MODE) Then
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        
        '//기본사항 - 부서 (사업장에 따른 부서변경)
        Set oCombo = oForm.Items("TeamCode").Specific

        If oCombo.ValidValues.Count > 0 Then
            For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                oCombo.ValidValues.Remove i, psk_Index
            Next i
            oCombo.ValidValues.Add "", ""
            oCombo.Select 0, psk_Index
        End If
        
        If oForm.DataSources.UserDataSources.Item("CLTCOD").ValueEx <> "" Then
            sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
            sQry = sQry & " WHERE Code = '1' AND U_Char2 = '" & oForm.Items.Item("CLTCOD").Specific.Value & "'"
            sQry = sQry & " ORDER BY U_Code"
            Call SetReDataCombo(oForm, sQry, oCombo)
            oCombo.ValidValues.Add "%", "전체"
            oCombo.Select "%", psk_ByValue
        End If
        
        '//담당 (사업장에 따른 담당변경)
    
        Set oCombo = oForm.Items("RspCode").Specific

        If oCombo.ValidValues.Count > 0 Then
            For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                oCombo.ValidValues.Remove i, psk_Index
            Next i
            oCombo.ValidValues.Add "", ""
            oCombo.Select 0, psk_Index
        End If
        
        If oForm.DataSources.UserDataSources.Item("CLTCOD").ValueEx <> "" Then
            sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
            sQry = sQry & " WHERE Code = '2' AND U_Char2 = '" & oForm.Items.Item("CLTCOD").Specific.Value & "'"
            sQry = sQry & " Order By U_Code"
            Call SetReDataCombo(oForm, sQry, oCombo)
            oCombo.ValidValues.Add "%", "전체"
            oCombo.Select "%", psk_ByValue
        End If
        
        tSeqAll = 0
        Call PH_PY104_DataLoad
        
        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", False)    '//문서추가
        
    ElseIf (oForm.Mode = fm_FIND_MODE) Then
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        
        '//부서
        Set oCombo = oForm.Items("TeamCode").Specific
        If oCombo.ValidValues.Count > 0 Then
            For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                oCombo.ValidValues.Remove i, psk_Index
            Next i
            oCombo.ValidValues.Add "", "-"
            oCombo.Select 0, psk_Index
        End If

        '//담당
        Set oCombo = oForm.Items("RspCode").Specific
        If oCombo.ValidValues.Count > 0 Then
            For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                oCombo.ValidValues.Remove i, psk_Index
            Next i
            oCombo.ValidValues.Add "", "-"
            oCombo.Select 0, psk_Index
        End If
        
        
        tSeqAll = 0
        oForm.Mode = fm_ADD_MODE
        Call PH_PY104_DataLoad
        
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
    
PH_PY104_FormItemEnabled_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY104_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
                    If PH_PY104_DataValidCheck = True Then
                        Call PH_PY104_DataFind
                    Else
                        BubbleEvent = False
                    End If
                End If
                If pval.ItemUID = "Btn_Save" Then
                    If PH_PY104_DataSave = False Then
                        BubbleEvent = False
                    End If
                End If
                If pval.ItemUID = "Btn_AllChk" Then
                    If oGrid2.Rows.Count > 0 Then
                         Call oForm.Freeze(True)
                        For i = 0 To oGrid2.Rows.Count - 1
                            Call oDS_PH_PY104_02.setValue("ChkBox", i, "Y")
                        Next i
                         Call oForm.Freeze(False)
                    End If
                End If
                If pval.ItemUID = "Btn_Copy" Then
                    Call PH_PY104_DataCopy
                End If
               
            ElseIf pval.BeforeAction = False Then

            End If
            
'----------------------------------------------------------
        Case et_KEY_DOWN: '//2
            If pval.BeforeAction = False Then
'                If pval.ItemUID = "HOBONGF" Or pval.ItemUID = "HOBONGT" Then
'                    If pval.CharPressed = "9" Then
'                        If oForm.Items(pval.ItemUID).Specific.Value = "" Then
'                            oForm.Items(pval.ItemUID).CLICK ct_Regular
'                            Sbo_Application.ActivateMenuItem ("7425")
'                        End If
'                    End If
'                End If
            End If
            
'----------------------------------------------------------
        Case et_GOT_FOCUS: '//3
            Select Case pval.ItemUID
            Case "Grid1"
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
                        '//기본사항 - 부서 (사업장에 따른 부서변경)
                        Set oCombo = oForm.Items("TeamCode").Specific
    
                        If oCombo.ValidValues.Count > 0 Then
                            For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                                oCombo.ValidValues.Remove i, psk_Index
                            Next i
                        End If
                        
                        sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
                        sQry = sQry & " WHERE Code = '1' AND U_Char2 = '" & oForm.Items.Item("CLTCOD").Specific.Value & "'"
                        sQry = sQry & " ORDER BY U_Code"
                        Call SetReDataCombo(oForm, sQry, oCombo)
                        oCombo.ValidValues.Add "%", "전체"
                        oCombo.Select "%", psk_ByValue
                        oForm.Items("TeamCode").DisplayDesc = True
                        
                        Call PH_PY104_DataLoad
                        
                        '//담당 (사업장에 따른 담당변경)
                    
                        Set oCombo = oForm.Items("RspCode").Specific
    
                        If oCombo.ValidValues.Count > 0 Then
                            For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                                oCombo.ValidValues.Remove i, psk_Index
                            Next i
                        End If
                        
                        sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
                        sQry = sQry & " WHERE Code = '2' AND U_Char2 = '" & oForm.Items.Item("CLTCOD").Specific.Value & "'"
                        sQry = sQry & " Order By U_Code"
                        Call SetReDataCombo(oForm, sQry, oCombo)
                        oCombo.ValidValues.Add "%", "전체"
                        oCombo.Select "%", psk_ByValue
                        oForm.Items("RspCode").DisplayDesc = True
                        
                    End If
                End If
            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_CLICK: '//6
           
            If pval.BeforeAction = True Then
                Select Case pval.ItemUID
                Case "Grid1"
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
                If pval.ActionSuccess = True Then
                    If pval.ItemUID = "Grid1" And pval.ColUID = "SLT" Then
                        Call Check_Seq(pval.ColUID, pval.Row)
                    End If
                End If
            End If
           
'----------------------------------------------------------
        Case et_DOUBLE_CLICK: '//7
            If pval.BeforeAction = True Then
                If pval.ItemUID = "Grid1" Then
                    oForm.Freeze (False)
                    BubbleEvent = False
                End If
                Exit Sub
            End If
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

                End If
            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_MATRIX_LOAD: '//11
            If pval.BeforeAction = True Then
            ElseIf pval.BeforeAction = False Then
                Call PH_PY104_FormItemEnabled
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
                Set oDS_PH_PY104_01 = Nothing
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
                Call PH_PY104_FormItemEnabled

            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//문서찾기
                Call PH_PY104_FormItemEnabled
            Case "1282":    '//문서추가
                Call PH_PY104_FormItemEnabled

            Case "1288", "1289", "1290", "1291":
                Call PH_PY104_FormItemEnabled
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
    Case "Grid1"
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

Sub PH_PY104_FormClear()
On Error GoTo PH_PY104_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY104'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.Value = 1
    Else
        oForm.Items("DocEntry").Specific.Value = DocEntry
    End If
    Exit Sub
PH_PY104_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY104_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY104_DataValidCheck() As Boolean
On Error GoTo PH_PY104_DataValidCheck_Error
    PH_PY104_DataValidCheck = False
    Dim i       As Long
    Dim j       As Long
    Dim sQry    As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    j = 0
    If oGrid1.Rows.Count > 0 Then
        For i = 0 To oGrid1.Rows.Count - 1
            If oGrid1.DataTable.GetValue("SLT", i) = "Y" Then
                If oGrid1.DataTable.GetValue("U_CSUCOD", i) <> "" Then
                    j = j + 1
                End If
            End If
        Next i
    Else
        Sbo_Application.SetStatusBarMessage "수당, 공제 테이블에 데이터가 없습니다.", bmt_Short, True
    End If
    
    If j = 0 Then
        Sbo_Application.SetStatusBarMessage "선택된 수당, 공제 데이터가 없습니다", bmt_Short, True
        PH_PY104_DataValidCheck = False
        Exit Function
    End If
    
    If oForm.Items("CLTCOD").Specific.Value = "" Then
        Sbo_Application.SetStatusBarMessage "사업장은 필수입니다.", bmt_Short, True
        oForm.Items("CLTCOD").CLICK ct_Regular
        PH_PY104_DataValidCheck = False
        Exit Function
    End If
    
    PH_PY104_DataValidCheck = True
    
        
    Exit Function
    
    Set oRecordSet = Nothing
    
    
PH_PY104_DataValidCheck_Error:
    Set oRecordSet = Nothing
    PH_PY104_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY104_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Function PH_PY104_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY104_Validate_Error
    PH_PY104_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY104A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.Value & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", bmt_Short, True
        PH_PY104_Validate = False
        GoTo PH_PY104_Validate_Exit
    End If
'
    If ValidateType = "수정" Then

    ElseIf ValidateType = "행삭제" Then

    ElseIf ValidateType = "취소" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY104_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY104_Validate_Error:
    PH_PY104_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY104_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Private Sub PH_PY104_DataFind()
    Dim i           As Long
    Dim iRow        As Long
    Dim sQry        As String
    Dim sQry2        As String
    
    Dim CLTCOD      As String
    Dim TeamCode    As String
    Dim RspCode     As String
    Dim PAYTYP      As String
    Dim JIGCODF     As String
    Dim JIGCODT     As String
    Dim HOBONGF     As String
    Dim HOBONGT     As String
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    Dim pRecordset  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY104_DataFind_Error

    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    Set pRecordset = oCompany.GetBusinessObject(BoRecordset)
    oForm.Freeze (True)
        
    '/ PH_PY104_TEMP 테이블 초기화
    sQry = "DELETE PH_PY104_TEMP"
    oRecordSet.DoQuery sQry
    '/ PH_PY104_TEMP2 테이블 초기화
    sQry = "DELETE PH_PY104_TEMP2"
    oRecordSet.DoQuery sQry
    
    '/ 그리드1 체크 데이터 PH_PY104_TEMP  저장
    If oGrid1.Rows.Count > 0 Then
        For i = 0 To oGrid1.Rows.Count - 1
            If oDS_PH_PY104_01.GetValue("SLT", i) = "Y" Then
                sQry = "EXEC PH_PY104_Grid1 '" & oDS_PH_PY104_01.GetValue("GBN", i) & "','"
                sQry = sQry & oDS_PH_PY104_01.GetValue("U_CSUCOD", i) & "','"
                sQry = sQry & oDS_PH_PY104_01.GetValue("U_CSUNAM", i) & "','"
                sQry = sQry & oDS_PH_PY104_01.GetValue("SEQ", i) & "'"
                oRecordSet.DoQuery sQry
            End If
        Next i
    End If
    
    CLTCOD = oForm.Items("CLTCOD").Specific.Value
    TeamCode = oForm.Items("TeamCode").Specific.Value
    RspCode = oForm.Items("RspCode").Specific.Value
    PAYTYP = oForm.Items("PAYTYP").Specific.Value
    JIGCODF = IIf(oForm.Items("JIGCODF").Specific.Value = "%", "00", oForm.Items("JIGCODF").Specific.Value)
    JIGCODT = IIf(oForm.Items("JIGCODT").Specific.Value = "%", "ZZ", oForm.Items("JIGCODT").Specific.Value)
    HOBONGF = IIf(oForm.Items("HOBONGF").Specific.Value = "", "000", oForm.Items("HOBONGF").Specific.Value)
    HOBONGT = IIf(oForm.Items("HOBONGT").Specific.Value = "", "ZZZ", oForm.Items("HOBONGT").Specific.Value)
    
    '/ 검색 조건 - 임시 테이블 저장 PH_PY104_TEMP2
    sQry = "Exec PH_PY104_Grid2 '" & CLTCOD & "','" & TeamCode & "',"
    sQry = sQry & "'" & RspCode & "','" & PAYTYP & "',"
    sQry = sQry & "'" & JIGCODF & "','" & JIGCODT & "',"
    sQry = sQry & "'" & HOBONGF & "','" & HOBONGT & "'"
    oRecordSet.DoQuery sQry
    
    '/ 그리드1 체크 데이터P PH_PY104_TEMP 불러옴
    sQry = "SELECT GUBUN, CSUCOD, CSUNAM FROM PH_PY104_TEMP ORDER BY SEQ"
    oRecordSet.DoQuery sQry
    
    '/ PH_PY104_TEMP2  데이터에 코드와 이름 붙여서 다시 셀렉트
    If oRecordSet.RecordCount > 0 Then
        sQry = "SELECT '' AS ChkBox , T0.CODE, T0.Name "
            For i = 1 To oRecordSet.RecordCount
                If Trim(oRecordSet.Fields(0).Value) = "수당" Then
                    sQry = sQry & ", ISNULL((SELECT U_FILD03 FROM [@PH_PY001B] WHERE U_FILD02 ='" & oRecordSet.Fields(2).Value & "'  AND Code = T0.Code ),0) AS N'" & Trim(oRecordSet.Fields(2).Value) & "'"
                    
                ElseIf Trim(oRecordSet.Fields(0).Value) = "공제" Then
                    sQry = sQry & ", ISNULL((SELECT U_FILD03 FROM [@PH_PY001C] WHERE U_FILD02 ='" & oRecordSet.Fields(2).Value & "'  AND Code = T0.Code ),0) AS N'" & Trim(oRecordSet.Fields(2).Value) & "'"
                End If
                oRecordSet.MoveNext
            Next i
        sQry = sQry & " FROM  PH_PY104_TEMP2 T0"
        Call oDS_PH_PY104_02.ExecuteQuery(sQry)
    End If
    

    iRow = oForm.DataSources.DataTables.Item(0).Rows.Count
    Call PH_PY104_TitleSetting_Grid2(iRow)
    
    Set oRecordSet = Nothing
    Set pRecordset = Nothing
    oForm.Freeze (False)
    Exit Sub
    
PH_PY104_DataFind_Error:
    Set oRecordSet = Nothing
    Set pRecordset = Nothing
    oForm.Freeze (False)
    Sbo_Application.SetStatusBarMessage "PH_PY104_DataFind_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub
Private Function PH_PY104_DataLoad()
    Dim i           As Long
    Dim sQry        As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY104_DataLoad_ERROR
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    oForm.Freeze (True)
    
    sQry = "SELECT '수당' AS GBN, T0.U_CSUCOD, T0.U_CSUNAM, '' AS SLT, SPACE(5) AS SEQ"
    sQry = sQry & " FROM [@PH_PY102B] T0 INNER JOIN [@PH_PY102A] T1 ON T0.Code = T1.Code"
    sQry = sQry & " WHERE U_FIXGBN = 'Y' AND T1.U_YM = (select Max(U_YM) AS U_YM from [@PH_PY102A])"
    sQry = sQry & " AND T1.U_CLTCOD = '" & oForm.Items("CLTCOD").Specific.Value & "'"
    sQry = sQry & " Union All"
    sQry = sQry & " SELECT '공제' AS GBN, T0.U_CSUCOD, T0.U_CSUNAM, '' AS SLT, SPACE(5) AS SEQ"
    sQry = sQry & " FROM [@PH_PY103B] T0 INNER JOIN [@PH_PY103A] T1 ON T0.Code = T1.Code"
    sQry = sQry & " WHERE U_FIXGBN = 'Y' AND T1.U_YM = (select Max(U_YM) AS U_YM from [@PH_PY103A])"
    sQry = sQry & " AND T1.U_CLTCOD = '" & oForm.Items("CLTCOD").Specific.Value & "'"

    oDS_PH_PY104_01.ExecuteQuery sQry
    
    Call PH_PY104_TitleSetting_Grid1
    
    
    oForm.Freeze (False)
    Exit Function
    
PH_PY104_DataLoad_ERROR:
    oForm.Freeze (False)
    Sbo_Application.SetStatusBarMessage "PH_PY104_DataLoad_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Function PH_PY104_DataCopy() As Boolean
    Dim i           As Long
    Dim j           As Long
    Dim ShiftDat    As String
    Dim First       As Integer
    Dim FirstData()
    Dim V_StatusCnt&, oProValue&, TOTCNT&, tRow&  '//progbar
    
    On Error GoTo PH_PY104_DataCopy_Error
    oForm.Freeze (True)
    
    PH_PY104_DataCopy = False
    
    If Not oProgBar Is Nothing Then
        oProgBar.Stop
        Set oProgBar = Nothing
    End If
    
    First = 0
    If oGrid2.Rows.Count > 0 Then
        '프로그레스 바    ///////////////////////////////////////
         Set oProgBar = Sbo_Application.StatusBar.CreateProgressBar("데이터 읽는중...!", 50, False)
         
         '최대값 구하기
         TOTCNT = oGrid2.Rows.Count
        
         V_StatusCnt = Round(TOTCNT / 50, 0)
         oProValue = 1
         tRow = 1
    '/////////////////////////////////////////////////////
    
        For i = 0 To oGrid2.Rows.Count - 1
            If oDS_PH_PY104_02.GetValue("ChkBox", i) = "Y" Then
                First = First + 1
                If First = 1 Then
                    ReDim FirstData(oGrid2.Columns.Count)
                    For j = 0 To oGrid2.Columns.Count - 1
                        FirstData(j) = oDS_PH_PY104_02.Columns.Item(j).Cells(i).Value
                    Next j
                Else
                    For j = 3 To oGrid2.Columns.Count - 1
                        oDS_PH_PY104_02.Columns.Item(j).Cells(i).Value = FirstData(j)
                    Next j
                End If
            End If
            
            tRow = tRow + 1
            If (TOTCNT > 50 And tRow = oProValue * V_StatusCnt) Or TOTCNT <= 50 Then
                oProgBar.Text = tRow & "/ " & TOTCNT & " 건 처리중...!"
                oProValue = oProValue + 1
                oProgBar.Value = oProValue
            End If
        Next i
    End If
    oProgBar.Stop
    
    If First = 0 Then
            Sbo_Application.SetStatusBarMessage "선택된 데이터가 없습니다.", bmt_Short, True
            PH_PY104_DataCopy = False
        Else
            Sbo_Application.StatusBar.SetText "선택된 필드의 전체 값 복사가 완료되었습니다.", bmt_Short, smt_Success
            PH_PY104_DataCopy = True
        End If
    
    Set oProgBar = Nothing
    
    oForm.Freeze (False)
    Exit Function
    
PH_PY104_DataCopy_Error:
    If Not oProgBar Is Nothing Then
        oProgBar.Stop
        Set oProgBar = Nothing
    End If
    Set oProgBar = Nothing
    oForm.Freeze (False)
    Sbo_Application.SetStatusBarMessage "PH_PY104_DataCopy_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Function PH_PY104_DataSave() As Boolean
    Dim i               As Long
    Dim j               As Long
    Dim SuDangLineId    As Long
    Dim GongjeLineId    As Long
    Dim sQry            As String
    Dim sQry2           As String
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    On Error GoTo PH_PY104_DataSave_Error
    
    PH_PY104_DataSave = False
  
    If oGrid2.Rows.Count > 0 Then
        For i = 0 To oGrid2.Rows.Count - 1
            If oDS_PH_PY104_02.GetValue("ChkBox", i) = "Y" Then
                For j = 3 To oGrid2.Columns.Count - 1
'                    oDS_PH_PY104_02.Columns(j).Cells(i).Value
'                    Trim (oDS_PH_PY104_02.GetValue("CODE", i))
                    sQry = "SELECT GUBUN FROM PH_PY104_TEMP WHERE CSUNAM = '" & oDS_PH_PY104_02.Columns(j).Name & "'"
                    oRecordSet.DoQuery sQry
                    
                    If oRecordSet.RecordCount > 0 Then
                        If Trim(oRecordSet.Fields(0).Value) = "수당" Then
                            sQry = "UPDATE [@PH_PY001B] SET U_FILD03 = '" & oDS_PH_PY104_02.Columns(j).Cells(i).Value & "'"
                            sQry = sQry & " WHERE Code = '" & Trim(oDS_PH_PY104_02.GetValue("CODE", i)) & "'"
                            sQry = sQry & " AND U_FILD02 = '" & oDS_PH_PY104_02.Columns(j).Name & "'"
                            oRecordSet.DoQuery sQry
                        ElseIf Trim(oRecordSet.Fields(0).Value) = "공제" Then
                            sQry = "UPDATE [@PH_PY001C] SET U_FILD03 = '" & oDS_PH_PY104_02.Columns(j).Cells(i).Value & "'"
                            sQry = sQry & " WHERE Code = '" & Trim(oDS_PH_PY104_02.GetValue("CODE", i)) & "'"
                            sQry = sQry & " AND U_FILD02 = '" & oDS_PH_PY104_02.Columns(j).Name & "'"
                            oRecordSet.DoQuery sQry
                        End If
                    End If
                Next j
                Sbo_Application.SetStatusBarMessage "[" & Trim(oDS_PH_PY104_02.GetValue("CODE", i)) & "] 의 수당,공제 데이터가 갱신중입니다..", bmt_Short, False
            End If
        Next i
        PH_PY104_DataSave = True
        Sbo_Application.SetStatusBarMessage "수당,공제 데이터가 갱신 되었습니다.", bmt_Short, False
    Else
        Sbo_Application.SetStatusBarMessage "데이터가 존재하지 않습니다.", bmt_Short, True
    End If
        

    Set oRecordSet = Nothing
    Exit Function
    
PH_PY104_DataSave_Error:
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "PH_PY104_DataSave_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Private Function PH_PY104_OHEM_DI_UPDATE() As Boolean
    Dim i                   As Long
    Dim j                   As Long
    Dim IErrCode            As Long
    Dim sErrMsg             As String
    Dim sQry                As String
    Dim sItemName           As String
    Dim EmpID               As String
    
    Dim oRecordSet          As SAPbobsCOM.Recordset
    
    Dim oOHEM               As SAPbobsCOM.EmployeesInfo

    On Error GoTo Error_Handle

    PH_PY104_OHEM_DI_UPDATE = False
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    Set oOHEM = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oEmployeesInfo)
    
    sQry = "SELECT U_Empid FROM [@PH_PY001A] WHERE Code = '" & Trim(oDS_PH_PY104_01.Columns.Item("Code").Cells(i).Value) & "'"
    
    oRecordSet.DoQuery sQry
    
    EmpID = oRecordSet.Fields(0).Value
    
    Sbo_Application.SetStatusBarMessage Trim(oDS_PH_PY104_01.Columns.Item("Code").Cells(i).Value) & "의 근무조가 갱신중입니다.", bmt_Short, False
    
    If EmpID = "" Then
        Sbo_Application.SetStatusBarMessage "empID 로드에 실패하였습니다.", bmt_Short, False
        GoTo Error_Handle
    End If
    If oOHEM.GetByKey(EmpID) = True Then
        
        With oOHEM
            .GetByKey (EmpID)
    
            .UserFields.Fields.Item("U_GNMUJO").Value = oDS_PH_PY104_01.Columns.Item("GNMUJO").Cells(i).Value      '//근무조
        
            If 0 <> oOHEM.Update Then
                oCompany.GetLastError IErrCode, sErrMsg
                Sbo_Application.SetStatusBarMessage sErrMsg, bmt_Short, True
                
            Else
                PH_PY104_OHEM_DI_UPDATE = True
                Sbo_Application.SetStatusBarMessage Trim(oDS_PH_PY104_01.Columns.Item("Code").Cells(i).Value) & "근무조가 갱신되었습니다.", bmt_Short, False
                
            End If
        End With
    End If
    
    Set oRecordSet = Nothing
    Set oOHEM = Nothing

    Exit Function

Error_Handle:
    Set oRecordSet = Nothing
    Set oOHEM = Nothing

End Function

Private Sub PH_PY104_TitleSetting_Grid1()
    Dim i               As Long
    Dim j               As Long
    Dim sQry            As String
     
    Dim COLNAM(4)       As String

    Dim oColumn         As SAPbouiCOM.EditTextColumn
    Dim oComboCol       As SAPbouiCOM.ComboBoxColumn
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
     
    On Error GoTo Error_Message
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    oForm.Freeze True
    

    COLNAM(0) = "구분"
    COLNAM(1) = "코드"
    COLNAM(2) = "코드명"
    COLNAM(3) = "선택"
    COLNAM(4) = "순서"

    For i = 0 To UBound(COLNAM)
        oGrid1.Columns.Item(i).TitleObject.Caption = COLNAM(i)
        If i >= 0 And i <= 2 Then
            oGrid1.Columns.Item(i).Editable = False
        ElseIf i = 3 Then
            oGrid1.Columns.Item(i).Editable = True
            oGrid1.Columns.Item(i).Type = gct_CheckBox
        ElseIf i = 4 Then
            oGrid1.Columns.Item(i).Editable = True
        End If
        
    Next i
    
    oGrid1.AutoResizeColumns

    oForm.Freeze False

    Set oColumn = Nothing

    Exit Sub

Error_Message:
    oForm.Freeze False
    Set oColumn = Nothing
    Sbo_Application.SetStatusBarMessage "PH_PY104_TitleSetting_Grid1 Error : " & Space(10) & Err.Description, bmt_Short, True
End Sub



Private Sub PH_PY104_TitleSetting_Grid2(iRow As Long)
    Dim i               As Long
    Dim j               As Long
    Dim sQry            As String
     
    Dim COLNAM(5)       As String

    Dim oColumn         As SAPbouiCOM.EditTextColumn
    Dim oComboCol       As SAPbouiCOM.ComboBoxColumn
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
     
    On Error GoTo Error_Message
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    oForm.Freeze True
    
    COLNAM(0) = "체크"
    COLNAM(1) = "사번"
    COLNAM(2) = "이름"
    
    
    For i = 0 To oGrid2.Columns.Count - 1
        If i = 0 Then
            oGrid2.Columns.Item(i).TitleObject.Caption = COLNAM(i)
            oGrid2.Columns.Item(i).Type = gct_CheckBox
            oGrid2.Columns.Item(i).Editable = True
        ElseIf i = 1 Or i = 2 Then
            oGrid2.Columns.Item(i).TitleObject.Caption = COLNAM(i)
            oGrid2.Columns.Item(i).Editable = False
        Else
            oGrid2.Columns.Item(i).Editable = True
            oGrid2.Columns.Item(i).RightJustified = True
        End If

    Next i
    
    oGrid2.AutoResizeColumns

    oForm.Freeze False

    Set oColumn = Nothing

    Exit Sub

Error_Message:
    oForm.Freeze False
    Set oColumn = Nothing
    Sbo_Application.SetStatusBarMessage "PH_PY104_TitleSetting_Grid2 Error : " & Space(10) & Err.Description, bmt_Short, True
End Sub


Private Sub Check_Seq(ColUID As String, Row As Long)
    Dim i       As Long
    Dim tSeq    As Long
    
     Call oForm.Freeze(True)
    If oGrid1.Rows.Count > 0 And Row >= 0 Then
        If tSeqAll < 0 Then tSeqAll = 0
        If oGrid1.DataTable.GetValue("SLT", Row) = "Y" Then
            tSeqAll = tSeqAll + 1
            oGrid1.DataTable.setValue "SEQ", Row, tSeqAll
        ElseIf oGrid1.DataTable.GetValue("SLT", Row) = "N" Then
            tSeqAll = tSeqAll - 1
            
            If oGrid1.DataTable.GetValue("SEQ", Row) = "" Then
                oGrid1.DataTable.setValue "SEQ", Row, 0
            End If
            tSeq = oGrid1.DataTable.GetValue("SEQ", Row)
            oGrid1.DataTable.setValue "SEQ", Row, ""
            
            For i = 0 To oGrid1.Rows.Count - 1
                If oGrid1.DataTable.GetValue("SEQ", i) > tSeq And oGrid1.DataTable.GetValue("SEQ", i) <> "" Then
                   oGrid1.DataTable.setValue "SEQ", i, CInt(oGrid1.DataTable.GetValue("SEQ", i)) - 1
                End If
            Next i
        End If
    End If
     Call oForm.Freeze(False)
End Sub
