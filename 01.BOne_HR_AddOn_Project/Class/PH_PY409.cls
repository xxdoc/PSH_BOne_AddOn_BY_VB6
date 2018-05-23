VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY409"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY409.cls
'//  Module         : 인사관리 > 연말정산관리
'//  Desc           : 정산기부금조정명세등록
'//********************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form

Public oGrid1           As SAPbouiCOM.Grid
Public oGrid2           As SAPbouiCOM.Grid
Public oMat             As SAPbouiCOM.Matrix
Public oDS_PH_PY409A    As SAPbouiCOM.DataTable
Public oDS_PH_PY409B    As SAPbouiCOM.DataTable
Private oDS_PH_PY409L   As SAPbouiCOM.DBDataSource

Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long

Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY409.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY409_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY409"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
'    oForm.DataBrowser.BrowseBy = "Code"
        
    oForm.PaneLevel = 1
    oForm.Freeze True
    Call PH_PY409_CreateItems
    Call PH_PY409_FormItemEnabled
    Call PH_PY409_EnableMenus
'    Call PH_PY409_SetDocument(oFromDocEntry01)
'    Call PH_PY409_FormResize
    
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

Private Function PH_PY409_CreateItems() As Boolean

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
    
    On Error GoTo PH_PY409_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
'    Set oDS_PH_PY409L = oForm.DataSources.DBDataSources("@PS_USERDS01")
    
    Set oGrid1 = oForm.Items("Grid01").Specific
    Set oGrid2 = oForm.Items("Grid02").Specific
    
    
    oForm.DataSources.DataTables.Add ("PH_PY409")
    oForm.DataSources.DataTables.Add ("PH_PY4091")
    
    
    oGrid1.DataTable = oForm.DataSources.DataTables.Item("PH_PY409")
    oGrid2.DataTable = oForm.DataSources.DataTables.Item("PH_PY4091")
    Set oDS_PH_PY409A = oForm.DataSources.DataTables.Item("PH_PY409")
    Set oDS_PH_PY409B = oForm.DataSources.DataTables.Item("PH_PY4091")
    

    '//----------------------------------------------------------------------------------------------
    '// 기본사항
    '//----------------------------------------------------------------------------------------------
    
    '//사업장
'    oCombo.Select 0, psk_Index
    oForm.Items("CLTCOD").DisplayDesc = True
    
    '//년도
    Call oForm.DataSources.UserDataSources.Add("Year", dt_SHORT_TEXT, 4)
    oForm.Items("Year").Specific.DataBind.SetBound True, "", "Year"
    
    '//사번
    Call oForm.DataSources.UserDataSources.Add("MSTCOD", dt_SHORT_TEXT, 10)
    oForm.Items("MSTCOD").Specific.DataBind.SetBound True, "", "MSTCOD"
    
    '//이름
    Call oForm.DataSources.UserDataSources.Add("FullName", dt_SHORT_TEXT, 20)
    oForm.Items("FullName").Specific.DataBind.SetBound True, "", "FullName"
    
    '//이월
    Call oForm.DataSources.UserDataSources.Add("dontew", dt_SUM)
    oForm.Items("dontew").Specific.DataBind.SetBound True, "", "dontew"
    
    '//세액공제대상
    Call oForm.DataSources.UserDataSources.Add("gongamt", dt_SUM)
    oForm.Items("gongamt").Specific.DataBind.SetBound True, "", "gongamt"
    
    '//합계
    Call oForm.DataSources.UserDataSources.Add("donttot", dt_SUM)
    oForm.Items("donttot").Specific.DataBind.SetBound True, "", "donttot"
    
    '//기부금코드  73
    Set oCombo = oForm.Items("gcode").Specific
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '73' AND U_UseYN= 'Y'"
    Call SetReDataCombo(oForm, sQry, oCombo, "Y")
    
    '//기부년도
    Call oForm.DataSources.UserDataSources.Add("gyyyy", dt_SHORT_TEXT, 4)
    oForm.Items("gyyyy").Specific.DataBind.SetBound True, "", "gyyyy"
    
    '//기부금액
    Call oForm.DataSources.UserDataSources.Add("gibuamt", dt_SUM)
    oForm.Items("gibuamt").Specific.DataBind.SetBound True, "", "gibuamt"
    
    '//전년까지공제된금액
    Call oForm.DataSources.UserDataSources.Add("jgamt", dt_SUM)
    oForm.Items("jgamt").Specific.DataBind.SetBound True, "", "jgamt"
    
    '//공제대상금액
    Call oForm.DataSources.UserDataSources.Add("gamt", dt_SUM)
    oForm.Items("gamt").Specific.DataBind.SetBound True, "", "gamt"
    
    '//해당년도공제금액
    Call oForm.DataSources.UserDataSources.Add("ygamt", dt_SUM)
    oForm.Items("ygamt").Specific.DataBind.SetBound True, "", "ygamt"
    
    '//소멸금액
    Call oForm.DataSources.UserDataSources.Add("disamt", dt_SUM)
    oForm.Items("disamt").Specific.DataBind.SetBound True, "", "disamt"
    
    '//이월금액
    Call oForm.DataSources.UserDataSources.Add("ewamt", dt_SUM)
    oForm.Items("ewamt").Specific.DataBind.SetBound True, "", "ewamt"
    
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
    
PH_PY409_CreateItems_Error:
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY409_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY409_EnableMenus()
    
    On Error GoTo PH_PY409_EnableMenus_Error

    Call oForm.EnableMenu("1283", False)     '//제거
    Call oForm.EnableMenu("1284", False)    '//취소
    Call oForm.EnableMenu("1293", False)     '//행삭제
    
    Exit Sub
    
PH_PY409_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY409_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY409_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY409_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY409_FormItemEnabled
'        Call PH_PY409_AddMatrixRow
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY409_FormItemEnabled
        oForm.Items("Code").Specific.VALUE = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY409_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY409_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY409_FormItemEnabled()
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim sQry        As String
    Dim i           As Long
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    
    Dim CLTCOD As String
    Dim sPosDate As String
    
    On Error GoTo PH_PY409_FormItemEnabled_Error
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Call oForm.Freeze(True)
    If (oForm.Mode = fm_ADD_MODE) Then
    
        
        Call oForm.EnableMenu("1281", False)     '//문서찾기
        Call oForm.EnableMenu("1282", True)    '//문서추가
        
        If Trim(oForm.Items("Year").Specific.VALUE) = "" Then
            oForm.Items("Year").Specific.VALUE = Format(Now, "YYYY") - 1
        End If
        
        If Trim(oForm.Items("MSTCOD").Specific.VALUE) = "" Then
            oForm.Items("MSTCOD").Specific.VALUE = ""
            oForm.Items("FullName").Specific.VALUE = ""
            oForm.Items("TeamName").Specific.VALUE = ""
            oForm.Items("RspName").Specific.VALUE = ""
            oForm.Items("ClsName").Specific.VALUE = ""
            oForm.Items("gongamt").Specific.VALUE = 0
        End If
        
        Set oCombo = oForm.Items("gcode").Specific
        oCombo.Select "0", psk_Index
        
        oForm.DataSources.UserDataSources.Item("gyyyy").VALUE = ""
        
        oForm.DataSources.UserDataSources.Item("gibuamt").VALUE = 0
        oForm.DataSources.UserDataSources.Item("jgamt").VALUE = 0
        oForm.DataSources.UserDataSources.Item("gamt").VALUE = 0
        oForm.DataSources.UserDataSources.Item("ygamt").VALUE = 0
        oForm.DataSources.UserDataSources.Item("disamt").VALUE = 0
        oForm.DataSources.UserDataSources.Item("ewamt").VALUE = 0
        
        
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
    
    '//Key set
    oForm.Items("CLTCOD").Enabled = True
    oForm.Items("Year").Enabled = True
    oForm.Items("MSTCOD").Enabled = True
    
    oForm.Items("gcode").Enabled = True
    oForm.Items("gyyyy").Enabled = True
    
    Set oCombo = Nothing
    Call oForm.Freeze(False)
    Exit Sub
    
PH_PY409_FormItemEnabled_Error:
    Set oCombo = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY409_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Public Sub Raise_FormItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
    Dim sQry        As String
    Dim i           As Long
    Dim tSex        As String
    Dim tBrith      As String
    Dim Day         As String
    Dim ActCode     As String
    Dim CLTCOD      As String
    Dim YEAR        As String
    Dim MSTCOD      As String
    Dim Div         As String
    Dim target      As String
    Dim relate      As String
    Dim YY          As String
    Dim Result      As String
    
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
                    If PH_PY409_DataValidCheck = False Then
                        BubbleEvent = False
                    End If
                End If
                
                If pval.ItemUID = "Btn_ret" Then
                    Call PH_PY409_MTX01
                End If
                
                                
                If pval.ItemUID = "Btn01" Then
                
'                        CLTCOD = oForm.Items("CLTCOD").Specific.VALUE
'                        YY = oForm.Items("Year").Specific.VALUE
'                        sQry = " Select CODE from [@PSH_PAYBACK] WHERE CODE = '" & YY & CLTCOD & "'"
'                        oRecordSet.DoQuery sQry
'
'                        Result = oRecordSet.Fields(0).VALUE
'                    If Result <> "" Then
'                        Call MDC_Com.MDC_GF_Message("등록불가 년도입니다. 담당자에게 문의바랍니다.", "E")
'                    End If
'                    If Result = "" Then
                        Call PH_PY409_SAVE
'                    End If
                End If
                                
                If pval.ItemUID = "Btn_del" Then
                    Call PH_PY409_Delete
                    Call PH_PY409_FormItemEnabled
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
                            Call PH_PY409_FormItemEnabled
                        End If
                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY409_FormItemEnabled
                        End If
                    ElseIf oForm.Mode = fm_OK_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY409_FormItemEnabled
                        End If
                    End If
'
                End Select
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
                    
                    If pval.ItemUID = "div" Then
                        If oForm.Items("div").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                                BubbleEvent = False
                        End If
                    End If
                    If pval.ItemUID = "target" Then
                        If oForm.Items("target").Specific.VALUE = "" Then
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
                    
                End If
            End If

            Call oForm.Freeze(False)
            Set oRecordSet = Nothing
'----------------------------------------------------------
        Case et_CLICK: '//6
            Call oForm.Freeze(True)
            If pval.BeforeAction = True Then
                Select Case pval.ItemUID
                Case "Grid01"
                    If pval.Row >= 0 Then
                        Select Case pval.ItemUID
                        Case "Grid01"
                            
                            
                            'Call oMat.SelectRow(pval.Row, True, False)
                            PH_PY409_MTX02 pval.ItemUID, pval.Row, pval.ColUID
                        End Select
                        
                    End If
                Case "Mat01"
                    If pval.Row >= 0 Then
                        Call oMat.SelectRow(pval.Row, True, False)
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
'                If pval.ItemUID = "Mat01" Then
'
'                    oForm.Items("kname").Specific.VALUE = oMat.Columns("kname").Cells(pval.Row).Specific.VALUE
'                    oForm.Items("juminno").Specific.VALUE = oMat.Columns("juminno").Cells(pval.Row).Specific.VALUE
'
'                End If
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
                    
                        Case "Year"
                            oForm.Items("MSTCOD").Specific.VALUE = ""
                            
                        Case "MSTCOD"
                            CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE)
                            YEAR = oForm.Items("Year").Specific.VALUE
                            MSTCOD = oForm.Items("MSTCOD").Specific.VALUE
                            
                            sQry = "Select Code,"
                            sQry = sQry + " FullName = U_FullName,"
                            sQry = sQry + " TeamName = Isnull((SELECT U_CodeNm"
                            sQry = sQry + " From [@PS_HR200L]"
                            sQry = sQry + " WHERE Code = '1'"
                            sQry = sQry + " And U_Code = U_TeamCode),''),"
                            sQry = sQry + " RspName  = Isnull((SELECT U_CodeNm"
                            sQry = sQry + " From [@PS_HR200L]"
                            sQry = sQry + " WHERE Code = '2'"
                            sQry = sQry + " And U_Code = U_RspCode),''),"
                            sQry = sQry + " ClsName  = Isnull((SELECT U_CodeNm"
                            sQry = sQry + " From [@PS_HR200L]"
                            sQry = sQry + " WHERE Code = '9'"
                            sQry = sQry + " And U_Code  = U_ClsCode"
                            sQry = sQry + " And U_Char3 = U_CLTCOD),'')"
                            sQry = sQry + " From [@PH_PY001A]"
                            sQry = sQry + " Where U_CLTCOD = '" + CLTCOD + "'"
                            sQry = sQry + " and Code = '" + MSTCOD + "'"
                              
                            oRecordSet.DoQuery sQry
                            
                            oForm.Items("FullName").Specific.VALUE = oRecordSet.Fields("FullName").VALUE
                            oForm.Items("TeamName").Specific.VALUE = oRecordSet.Fields("TeamName").VALUE
                            oForm.Items("RspName").Specific.VALUE = oRecordSet.Fields("RspName").VALUE
                            oForm.Items("ClsName").Specific.VALUE = oRecordSet.Fields("ClsName").VALUE
                            
                            
                            '정산금액 찿기
                            sQry = "Select dont = dont_t + dont1_t + dont2_t + dont3_t + Isnull(poldont_t,0), "
                            sQry = sQry + " dontew = dontew "
                            sQry = sQry + " From [p_seoycpt] "
                            sQry = sQry + " WHERE saup = '" + CLTCOD + "'"
                            sQry = sQry + " and yyyy = '" + YEAR + "'"
                            sQry = sQry + " and sabun = '" + MSTCOD + "'"
                              
                            oRecordSet.DoQuery sQry
                            
                            oForm.Items("dontew").Specific.VALUE = oRecordSet.Fields("dontew").VALUE
                            oForm.Items("gongamt").Specific.VALUE = oRecordSet.Fields("dont").VALUE
                            oForm.Items("donttot").Specific.VALUE = oRecordSet.Fields("dont").VALUE + oRecordSet.Fields("dontew").VALUE
                            
                            '기부자료집계 표시
                            sQry = "EXEC PH_PY409_03 '" & CLTCOD & "', '" & YEAR & "', '" & MSTCOD & "'"
                            Call oDS_PH_PY409B.ExecuteQuery(sQry)
                            
                        
                    End Select
                 
                End If
            End If
'            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_MATRIX_LOAD: '//11
            If pval.BeforeAction = True Then
            ElseIf pval.BeforeAction = False Then
'                oMat.LoadFromDataSource
'                Call PH_PY409_AddMatrixRow
                
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
                Set oDS_PH_PY409A = Nothing
                
'                Set oMat = Nothing
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
'                    Call MDC_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY409A", "Code")
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
'                Call PH_PY409_FormItemEnabled
        End Select
    ElseIf (pval.BeforeAction = False) Then
        Select Case pval.MenuUID
            Case "1283":
                oForm.Mode = fm_ADD_MODE
                Call PH_PY409_FormItemEnabled
'                Call PH_PY409_AddMatrixRow
            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//문서찾기
                Call PH_PY409_FormItemEnabled
'                Call PH_PY409_AddMatrixRow
                oForm.Items("Code").CLICK ct_Regular
            Case "1282":    '//문서추가
                Call PH_PY409_FormItemEnabled
'                Call PH_PY409_AddMatrixRow
            Case "1288", "1289", "1290", "1291":
                Call PH_PY409_FormItemEnabled
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


Sub PH_PY409_FormClear()
On Error GoTo PH_PY409_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY409'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PH_PY409_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY409_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY409_DataValidCheck() As Boolean
On Error GoTo PH_PY409_DataValidCheck_Error
    PH_PY409_DataValidCheck = False
    Dim i       As Long
    Dim j       As Long
    
    Dim sQry    As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Exit Function
    
    Set oRecordSet = Nothing
    
    
PH_PY409_DataValidCheck_Error:
    Set oRecordSet = Nothing
    PH_PY409_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY409_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY409_MTX01()

    '//메트릭스에 데이터 로드
    
    Dim i       As Long
    Dim sQry    As String
    Dim iRow        As Long
    
    Dim Param01 As String
    Dim Param02 As String
    Dim Param03 As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY409_MTX01_Error

    Call oForm.Freeze(True)
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    Param01 = Trim(oForm.Items("CLTCOD").Specific.VALUE)
    Param02 = oForm.Items("Year").Specific.VALUE
    Param03 = oForm.Items("MSTCOD").Specific.VALUE
        
    If Trim(Param01) = "" Then
        Call MDC_Com.MDC_GF_Message("사업장이 없습니다. 확인바랍니다..", "E")
        GoTo PH_PY409_MTX01_Exit
    End If
        
    If Trim(Param02) = "" Then
        Call MDC_Com.MDC_GF_Message("년도가 없습니다. 확인바랍니다..", "E")
        GoTo PH_PY409_MTX01_Exit
    End If
        
    If Trim(Param03) = "" Then
        Call MDC_Com.MDC_GF_Message("사번이 없습니다. 확인바랍니다..", "E")
        GoTo PH_PY409_MTX01_Exit
    End If
        
        
        
    sQry = "EXEC PH_PY409_01 '" & Param01 & "', '" & Param02 & "', '" & Param03 & "'"

    Call oDS_PH_PY409A.ExecuteQuery(sQry)
    
    
    
    iRow = oForm.DataSources.DataTables.Item(0).Rows.Count
    
    'Call PH_PY409_TitleSetting(iRow)
    
    oForm.Update
    
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Sub
PH_PY409_MTX01_Exit:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Sub
PH_PY409_MTX01_Error:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY409_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub
Private Sub PH_PY409_MTX02(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)


    '//그리드 자료를 head에 로드
    
    Dim i       As Long
    Dim sQry    As String
    Dim sRow As Long
    
    Dim Param01 As String
    Dim Param02 As String
    Dim Param03 As String
    Dim Param04 As String
    Dim Param05 As String
    
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY409_MTX02_Error

    Call oForm.Freeze(True)
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    sRow = oRow
    
    
    Param01 = Trim(oForm.Items("CLTCOD").Specific.VALUE)
    Param02 = oDS_PH_PY409A.Columns.Item("연도").Cells(oRow).VALUE
    Param03 = oDS_PH_PY409A.Columns.Item("사번").Cells(oRow).VALUE
    Param04 = oDS_PH_PY409A.Columns.Item("기부금코드").Cells(oRow).VALUE
    Param05 = oDS_PH_PY409A.Columns.Item("기부년도").Cells(oRow).VALUE
    
    
    sQry = "EXEC PH_PY409_02 '" & Param01 & "', '" & Param02 & "', '" & Param03 & "', '" & Param04 & "', '" & Param05 & "'"
    Call oRecordSet.DoQuery(sQry)
    
    If (oRecordSet.RecordCount = 0) Then
        
      '  oForm.Items("MSTCOD").Specific.VALUE = oDS_PH_PY409A.Columns.Item("MSTCOD").Cells(oRow).VALUE
      '  oForm.Items("FullName").Specific.VALUE = oDS_PH_PY409A.Columns.Item("FullName").Cells(oRow).VALUE
        
        Set oCombo = oForm.Items("gcode").Specific
        oCombo.Select "0", psk_Index
        
        oForm.DataSources.UserDataSources.Item("gyyyy").VALUE = ""
        
        oForm.DataSources.UserDataSources.Item("gibuamt").VALUE = 0
        oForm.DataSources.UserDataSources.Item("jgamt").VALUE = 0
        oForm.DataSources.UserDataSources.Item("gamt").VALUE = 0
        oForm.DataSources.UserDataSources.Item("ygamt").VALUE = 0
        oForm.DataSources.UserDataSources.Item("disamt").VALUE = 0
        oForm.DataSources.UserDataSources.Item("ewamt").VALUE = 0
        
        'oForm.Items("TeamName").Specific.VALUE = ""
        'oForm.Items("RspName").Specific.VALUE = ""
        'oForm.Items("ClsName").Specific.VALUE = ""
    
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
        GoTo PH_PY409_MTX02_Exit
    End If
    
    'oForm.Items("Year").Specific.VALUE = oRecordSet.Fields("Year").VALUE
    'oForm.Items("MSTCOD").Specific.VALUE = oRecordSet.Fields("MSTCOD").VALUE
    'oForm.Items("FullName").Specific.VALUE = oRecordSet.Fields("FullName").VALUE
    
'    '//부서
    'oForm.Items("TeamName").Specific.VALUE = oRecordSet.Fields("TeamName").VALUE
    'oForm.Items("RspName").Specific.VALUE = oRecordSet.Fields("RspName").VALUE
    'oForm.Items("ClsName").Specific.VALUE = oRecordSet.Fields("ClsName").VALUE
    
    Set oCombo = oForm.Items("gcode").Specific
    oCombo.Select oRecordSet.Fields("gcode").VALUE, psk_ByValue
    
    
    oForm.DataSources.UserDataSources.Item("gyyyy").VALUE = oRecordSet.Fields("gyyyy").VALUE
    
    oForm.DataSources.UserDataSources.Item("gibuamt").VALUE = oRecordSet.Fields("gibuamt").VALUE
    oForm.DataSources.UserDataSources.Item("jgamt").VALUE = oRecordSet.Fields("jgamt").VALUE
    oForm.DataSources.UserDataSources.Item("gamt").VALUE = oRecordSet.Fields("gamt").VALUE
    oForm.DataSources.UserDataSources.Item("ygamt").VALUE = oRecordSet.Fields("ygamt").VALUE
    oForm.DataSources.UserDataSources.Item("disamt").VALUE = oRecordSet.Fields("disamt").VALUE
    oForm.DataSources.UserDataSources.Item("ewamt").VALUE = oRecordSet.Fields("ewamt").VALUE
    
    oForm.Update
    
    oForm.ActiveItem = "gibuamt"
    
    '//key set
    oForm.Items("CLTCOD").Enabled = False
    oForm.Items("Year").Enabled = False
    oForm.Items("MSTCOD").Enabled = False
    
    
    oForm.Items("gcode").Enabled = False
    oForm.Items("gyyyy").Enabled = False
    
    
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    
    Exit Sub
PH_PY409_MTX02_Exit:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    
    Exit Sub
PH_PY409_MTX02_Error:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY409_MTX02_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY409_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY409_Validate_Error
    PH_PY409_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY409A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", bmt_Short, True
        PH_PY409_Validate = False
        GoTo PH_PY409_Validate_Exit
    End If
'
    If ValidateType = "수정" Then

    ElseIf ValidateType = "행삭제" Then

    ElseIf ValidateType = "취소" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY409_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY409_Validate_Error:
    PH_PY409_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY409_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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


Private Sub PH_PY409_SAVE()

    '//데이타 저장
    
    Dim i       As Long
    Dim sQry    As String
    
    Dim saup$, sabun$, FullName$, yyyy$
    Dim gcode$, gyyyy$
    
    Dim gibuamt, jgamt, gamt, ygamt, disamt, ewamt As Double
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY409_SAVE_Error
    
    saup = Trim(oForm.Items("CLTCOD").Specific.VALUE)
    yyyy = oForm.Items("Year").Specific.VALUE
    sabun = oForm.Items("MSTCOD").Specific.VALUE
    gcode = oForm.Items("gcode").Specific.VALUE
    gyyyy = oForm.Items("gyyyy").Specific.VALUE
    gibuamt = oForm.Items("gibuamt").Specific.VALUE
    jgamt = oForm.Items("jgamt").Specific.VALUE
    gamt = oForm.Items("gamt").Specific.VALUE
    ygamt = oForm.Items("ygamt").Specific.VALUE
    disamt = oForm.Items("disamt").Specific.VALUE
    ewamt = oForm.Items("ewamt").Specific.VALUE
    
    FullName = oForm.Items("FullName").Specific.VALUE

    Call oForm.Freeze(True)
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    
    If Trim(yyyy) = "" Then
        Call MDC_Com.MDC_GF_Message("년도가 없습니다. 확인바랍니다..", "E")
        GoTo PH_PY409_SAVE_Exit
    End If
    
    If Trim(saup) = "" Then
        Call MDC_Com.MDC_GF_Message("사업장이 없습니다. 확인바랍니다..", "E")
        GoTo PH_PY409_SAVE_Exit
    End If
    If Trim(sabun) = "" Then
        Call MDC_Com.MDC_GF_Message("사번이 없습니다. 확인바랍니다..", "E")
        GoTo PH_PY409_SAVE_Exit
    End If
      
    sQry = " Select Count(*) From [p_seoygibucont] Where saup = '" & saup & "' And yyyy = '" & yyyy & "' And sabun = '" & sabun & "'"
    sQry = sQry + " And gcode = '" & gcode & "' And gyyyy = '" & gyyyy & "'"
    oRecordSet.DoQuery sQry
    
    If oRecordSet.Fields(0).VALUE > 0 Then
        '//갱신
        sQry = "Update [p_seoygibucont] set "
        sQry = sQry + "gibuamt = " & gibuamt & ","
        sQry = sQry + "jgamt = " & jgamt & ","
        sQry = sQry + "gamt = " & gamt & ","
        sQry = sQry + "ygamt = " & ygamt & ","
        sQry = sQry + "disamt = " & disamt & ","
        sQry = sQry + "ewamt = " & ewamt & ""
        
        sQry = sQry + " Where saup = '" & saup & "' And yyyy = '" & yyyy & "' And sabun = '" & sabun & "'"
        sQry = sQry + " And gcode = '" & gcode & "' And gyyyy = '" & gyyyy & "'"
        
        oRecordSet.DoQuery sQry
        
    Else
        
        '//신규
        sQry = "INSERT INTO [p_seoygibucont]"
        sQry = sQry & " ("
        sQry = sQry & "saup,"
        sQry = sQry & "yyyy,"
        sQry = sQry & "sabun,"
        sQry = sQry & "gcode,"
        sQry = sQry & "gyyyy,"
        sQry = sQry & "gibuamt,"
        sQry = sQry & "jgamt,"
        sQry = sQry & "gamt,"
        sQry = sQry & "ygamt,"
        sQry = sQry & "disamt,"
        sQry = sQry & "ewamt)"
        
        sQry = sQry & " VALUES("

        sQry = sQry & "'" & saup & "',"
        sQry = sQry & "'" & yyyy & "',"
        sQry = sQry & "'" & sabun & "',"
        sQry = sQry & "'" & gcode & "',"
        sQry = sQry & "'" & gyyyy & "',"

        sQry = sQry & gibuamt & ","
        sQry = sQry & jgamt & ","
        sQry = sQry & gamt & ","
        sQry = sQry & ygamt & ","
        sQry = sQry & disamt & ","
        sQry = sQry & ewamt & " )"

        oRecordSet.DoQuery sQry
    End If
    
    Call PH_PY409_FormItemEnabled
    
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)

    Call PH_PY409_MTX01

    Exit Sub
    
PH_PY409_SAVE_Exit:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    
    Exit Sub
PH_PY409_SAVE_Error:
    Call oForm.Freeze(False)
   
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "PH_PY409_SAVE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub PH_PY409_Delete()
    '//선택된 자료 삭제
    Dim i       As Long
    Dim sQry    As String
    Dim cnt     As Integer
    
    Dim saup$, sabun$, FullName$, yyyy$
    Dim gcode$, gyyyy$
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY409_Delete_Error
    
    saup = Trim(oForm.Items("CLTCOD").Specific.VALUE)
    yyyy = oForm.Items("Year").Specific.VALUE
    sabun = oForm.Items("MSTCOD").Specific.VALUE
    gcode = oForm.Items("gcode").Specific.VALUE
    gyyyy = oForm.Items("gyyyy").Specific.VALUE
    
    FullName = oForm.Items("gcode").Specific.VALUE
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
      
    sQry = " Select Count(*) From [p_seoygibucont] Where saup = '" & saup & "' And yyyy = '" & yyyy & "' And sabun = '" & sabun & "'"
    sQry = sQry + " And gcode = '" & gcode & "' And gyyyy = '" & gyyyy & "'"
        
    oRecordSet.DoQuery sQry
    
    cnt = oRecordSet.Fields(0).VALUE
    If cnt > 0 Then
    
        If Trim(yyyy) = "" Then
            Call MDC_Com.MDC_GF_Message("년도가 없습니다. 확인바랍니다..", "E")
            GoTo PH_PY409_Delete_Exit
        End If
        
        If Trim(saup) = "" Then
            Call MDC_Com.MDC_GF_Message("사업장이 없습니다. 확인바랍니다..", "E")
            GoTo PH_PY409_Delete_Exit
        End If
        If Trim(sabun) = "" Then
            Call MDC_Com.MDC_GF_Message("사번이 없습니다. 확인바랍니다..", "E")
            GoTo PH_PY409_Delete_Exit
        End If
        
        
        If Sbo_Application.MessageBox(" 선택한대상자('" & FullName & "')을 삭제하시겠습니까? ?", "2", "예", "아니오") = "1" Then
            sQry = "Delete From [p_seoygibucont] Where saup = '" & saup & "' And yyyy = '" & yyyy & "' And sabun = '" & sabun & "'"
            sQry = sQry + " And gcode = '" & gcode & "' And gyyyy = '" & gyyyy & "'"
            oRecordSet.DoQuery sQry
        End If
    End If
    
   
    Call oForm.Freeze(False)
    
        
    Call PH_PY409_MTX01
    
    Set oRecordSet = Nothing
    
    
    Exit Sub
PH_PY409_Delete_Exit:
    Set oRecordSet = Nothing
        
    Call oForm.Freeze(False)
    Exit Sub
PH_PY409_Delete_Error:
    Set oRecordSet = Nothing
   
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY409_Delete_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub PH_PY409_TitleSetting(iRow As Long)
'    Dim i               As Long
'    Dim j               As Long
'    Dim sQry            As String
'
'    Dim COLNAM(12)       As String
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
'    COLNAM(0) = "년도"
'    COLNAM(1) = "사번"
'    COLNAM(2) = "공제구분코드"
'    COLNAM(3) = "공제구분"
'    COLNAM(4) = "공제대상코드"
'    COLNAM(5) = "공제대상"
'    COLNAM(6) = "관계코드"
'    COLNAM(7) = "관계"
'    COLNAM(8) = "대상자성명"
'    COLNAM(9) = "주민번호"
'    COLNAM(10) = "금액(국세청)"
'    COLNAM(11) = "금액(국세청외)"
'    COLNAM(12) = "합계금액"
'
'    For i = 0 To UBound(COLNAM)
'        oGrid1.Columns.Item(i).TitleObject.Caption = COLNAM(i)
'        oGrid1.Columns.Item(i).Editable = False
'        If COLNAM(i) = "사번" Or COLNAM(i) = "공제구분코드" Or COLNAM(i) = "공제대상코드" Or COLNAM(i) = "관계코드" Or COLNAM(i) = "주민번호" Then
'            oGrid1.Columns.Item(i).Visible = False
'        End If
'
'        oGrid1.Columns.Item(i).RightJustified = True
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
'
'Error_Message:
'    oForm.Freeze False
'    Set oColumn = Nothing
'    Sbo_Application.SetStatusBarMessage "PH_PY409_TitleSetting Error : " & Space(10) & Err.Description, bmt_Short, True
End Sub











