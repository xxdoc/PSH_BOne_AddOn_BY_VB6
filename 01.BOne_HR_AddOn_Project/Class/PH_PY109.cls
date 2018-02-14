VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY109"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY109.cls
'//  Module         : 인사관리 > 급여관리
'//  Desc           : 급상여변동자료등록
'//********************************************************************************
Option Explicit

Public oFormUniqueID            As String
Public oForm                    As SAPbouiCOM.Form

Public oMat1                    As SAPbouiCOM.Matrix
Public oMat2                    As SAPbouiCOM.Matrix
Private oDS_PH_PY109A           As SAPbouiCOM.DBDataSource
Private oDS_PH_PY109B           As SAPbouiCOM.DBDataSource
Private oDS_PH_PY109Z           As SAPbouiCOM.DBDataSource

Private oLastItemUID            As String
Private oLastColUID             As String
Private oLastColRow             As Long

Private sCode       As String
Private sCLTCOD     As String
Private sYM         As String
Private sJOBTYP     As String
Private sJOBGBN     As String
Private sJOBTRG     As String

Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY109.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    
    oFormUniqueID = "PH_PY109_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY109"
    
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
    oForm.DataBrowser.BrowseBy = "Code"
        
    oForm.Freeze True
    Call PH_PY109_CreateItems
    Call PH_PY109_EnableMenus
    Call PH_PY109_SetDocument(oFromDocEntry01)
'    Call PH_PY109_FormResize
    
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

Private Function PH_PY109_CreateItems() As Boolean

    Dim sQry        As String
    Dim i           As Long
    
    Dim oCheck      As SAPbouiCOM.CheckBox
    Dim oEdit       As SAPbouiCOM.EditText
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns    As SAPbouiCOM.Columns
    Dim optBtn      As SAPbouiCOM.OptionBtn
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY109_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '/Matrix
    Set oDS_PH_PY109A = oForm.DataSources.DBDataSources("@PH_PY109A")   '//헤더
    Set oDS_PH_PY109B = oForm.DataSources.DBDataSources("@PH_PY109B")   '//라인
    Set oDS_PH_PY109Z = oForm.DataSources.DBDataSources("@PH_PY109Z")   '//라인

    Set oMat1 = oForm.Items("Mat1").Specific
    Set oMat2 = oForm.Items("Mat2").Specific       '
    
    Call oForm.DataSources.UserDataSources.Add("S_Amt01", dt_SUM)
    oForm.Items("S_Amt01").Specific.DataBind.SetBound True, "", "S_Amt01"
    
    Call oForm.DataSources.UserDataSources.Add("S_Amt02", dt_SUM)
    oForm.Items("S_Amt02").Specific.DataBind.SetBound True, "", "S_Amt02"
    
    Call oForm.DataSources.UserDataSources.Add("S_Amt03", dt_SUM)
    oForm.Items("S_Amt03").Specific.DataBind.SetBound True, "", "S_Amt03"
    
    Call oForm.DataSources.UserDataSources.Add("S_Amt04", dt_SUM)
    oForm.Items("S_Amt04").Specific.DataBind.SetBound True, "", "S_Amt04"
    
    Call oForm.DataSources.UserDataSources.Add("S_Amt05", dt_SUM)
    oForm.Items("S_Amt05").Specific.DataBind.SetBound True, "", "S_Amt05"
    
    Call oForm.DataSources.UserDataSources.Add("S_Amt06", dt_SUM)
    oForm.Items("S_Amt06").Specific.DataBind.SetBound True, "", "S_Amt06"
    
    Call oForm.DataSources.UserDataSources.Add("S_Amt07", dt_SUM)
    oForm.Items("S_Amt07").Specific.DataBind.SetBound True, "", "S_Amt07"
    
    Call oForm.DataSources.UserDataSources.Add("S_Amt08", dt_SUM)
    oForm.Items("S_Amt08").Specific.DataBind.SetBound True, "", "S_Amt08"
    
    Call oForm.DataSources.UserDataSources.Add("S_Amt09", dt_SUM)
    oForm.Items("S_Amt09").Specific.DataBind.SetBound True, "", "S_Amt09"
    
    Call oForm.DataSources.UserDataSources.Add("S_Amt10", dt_SUM)
    oForm.Items("S_Amt10").Specific.DataBind.SetBound True, "", "S_Amt10"
    
    Call oForm.DataSources.UserDataSources.Add("S_Amt11", dt_SUM)
    oForm.Items("S_Amt11").Specific.DataBind.SetBound True, "", "S_Amt11"
    
    Call oForm.DataSources.UserDataSources.Add("S_Amt12", dt_SUM)
    oForm.Items("S_Amt12").Specific.DataBind.SetBound True, "", "S_Amt12"
    

    
    
    oMat1.SelectionMode = ms_NotSupported
    oMat1.AutoResizeColumns
    oMat1.SelectionMode = ms_NotSupported
    oMat1.AutoResizeColumns
    
    
    '//----------------------------------------------------------------------------------------------
    '// 아이템 설정
    '//----------------------------------------------------------------------------------------------
     '//사업장
    Set oCombo = oForm.Items("CLTCOD").Specific
'    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
'    Call SetReDataCombo(oForm, sQry, oCombo)
    oForm.Items("CLTCOD").DisplayDesc = True
    
    '/ 귀속년월
     oForm.Items("YM").Specific.VALUE = Format$(Now, "YYYYMM")
     
    '// 지급종류
    Set oCombo = oForm.Items("JOBTYP").Specific
    oCombo.ValidValues.Add "1", "급여"
    oCombo.ValidValues.Add "2", "상여"
    oForm.Items("JOBTYP").DisplayDesc = True
    
    '// 지급구분
    Set oCombo = oForm.Items("JOBGBN").Specific
    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P212' AND U_UseYN= 'Y'"
    Call SetReDataCombo(oForm, sQry, oCombo)
    oForm.Items("JOBGBN").DisplayDesc = True
    
    '// 지급대상
    Set oCombo = oForm.Items("JOBTRG").Specific
    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P213' AND U_UseYN= 'Y'"
    Call SetReDataCombo(oForm, sQry, oCombo)
    oForm.Items("JOBTRG").DisplayDesc = True
    
    '// 라인-------------------------------------------------------------------------------------------
    '//사번

'    '// 부서명
'    Set oColumn = oMat1.Columns("DPTNAM")
'    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='1' AND U_UseYN= 'Y'"
'    oRecordSet.DoQuery sQry
'    If oRecordSet.RecordCount > 0 Then
'        Do Until oRecordSet.EOF
'            oColumn.ValidValues.Add Trim$(oRecordSet.Fields(0).Value), Trim$(oRecordSet.Fields(1).Value)
'            oRecordSet.MoveNext
'        Loop
'    End If
'    oColumn.DisplayDesc = True
    
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
    
PH_PY109_CreateItems_Error:
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY109_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Private Sub PH_PY109_EnableMenus()
    
    On Error GoTo PH_PY109_EnableMenus_Error

    Call oForm.EnableMenu("1283", True)     '//제거
    Call oForm.EnableMenu("1284", False)    '//취소
    Call oForm.EnableMenu("1293", True)     '//행삭제
    
    Exit Sub
    
PH_PY109_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY109_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY109_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY109_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY109_FormItemEnabled
        Call PH_PY109_AddMatrixRow
        Call PH_PY109_TitleSetting_Matrix
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY109_FormItemEnabled
        oForm.Items("Code").Specific.VALUE = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY109_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY109_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY109_FormItemEnabled()
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oConditions As SAPbouiCOM.Conditions
    
    On Error GoTo PH_PY109_FormItemEnabled_Error
    
    
    Call oForm.Freeze(True)
    If (oForm.Mode = fm_ADD_MODE) Then
        oForm.Items("CLTCOD").Enabled = True
        oForm.Items("YM").Enabled = True
        oForm.Items("JOBTYP").Enabled = True
        oForm.Items("JOBGBN").Enabled = True
        oForm.Items("JOBTRG").Enabled = True
        
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        
        '/ 귀속년월
        oForm.Items("YM").Specific.VALUE = Format$(Now, "YYYYMM")
        
        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", False)    '//문서추가
        
        Call PH_PY109_AddMatrixRow
    ElseIf (oForm.Mode = fm_FIND_MODE) Then
        oForm.Items("CLTCOD").Enabled = True
        oForm.Items("YM").Enabled = True
        oForm.Items("JOBTYP").Enabled = True
        oForm.Items("JOBGBN").Enabled = True
        oForm.Items("JOBTRG").Enabled = True
        
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        
        Call oForm.EnableMenu("1281", False)    '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가
    ElseIf (oForm.Mode = fm_OK_MODE) Then
        oForm.Items("CLTCOD").Enabled = False
        oForm.Items("YM").Enabled = False
        oForm.Items("JOBTYP").Enabled = False
        oForm.Items("JOBGBN").Enabled = False
        oForm.Items("JOBTRG").Enabled = False
        
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD", False)
        
        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가
    End If
    Call oForm.Freeze(False)
    Exit Sub
    
PH_PY109_FormItemEnabled_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY109_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Public Sub Raise_FormItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
    Dim sQry        As String
    Dim i           As Long
    Dim FullName    As String
    Dim FindYN      As String
    
    
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    
    Dim oConditions As SAPbouiCOM.Conditions
    
    On Error GoTo Raise_FormItemEvent_Error
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Select Case pval.EventType
        Case et_ITEM_PRESSED: '//1
            If pval.BeforeAction = True Then
                If pval.ItemUID = "1" Then
                    If oForm.Mode = fm_ADD_MODE Or oForm.Mode = fm_UPDATE_MODE Then
                        If PH_PY109_DataValidCheck(pval.ItemUID) = False Then
                            BubbleEvent = False
                        End If
                    ElseIf oForm.Mode = fm_FIND_MODE Then
                    End If
                End If
                If pval.ItemUID = "Btn_Set" Then
                    If PH_PY109_DataValidCheck(pval.ItemUID) = False Then
                        BubbleEvent = False
                    End If
                End If
            ElseIf pval.BeforeAction = False Then
                If pval.ItemUID = "1" Then
                    If pval.ActionSuccess = True Then
                        Call PH_PY109_FormItemEnabled
                        Call PH_PY109_TitleSetting_Matrix
                    End If
                End If
                If pval.ItemUID = "Btn_Upload" Then
                    oMat1.FlushToDataSource
                    Call PH_PY109_Excel_Upload
                    Call PH_PY109_AddMatrixRow
                    
                    Call Sum_Display
                    
                End If
                If pval.ItemUID = "Btn_Set" Then
                    If pval.ActionSuccess = True Then
                        Call PH_PY109_LoadData_SudangGongje
                    End If
                End If
                
                If pval.ItemUID = "Search" Then
                    FindYN = "N"
                    FullName = oForm.Items("FullName").Specific.VALUE
                    For i = 1 To oMat1.VisualRowCount - 1
                        If oMat1.Columns("MSTNAM").Cells(i).Specific.VALUE = FullName Then
                            FindYN = "Y"
                            Call oMat1.SelectRow(i, True, False)
                        End If
                    Next
                    If FindYN = "N" Then
                        Sbo_Application.SetStatusBarMessage "찾는 사원이 없습니다.", bmt_Short, True
                    End If
                    
                End If
            End If
     
'----------------------------------------------------------
        Case et_KEY_DOWN: '//2
             If pval.BeforeAction = True And pval.ItemUID = "Mat1" And pval.ColUID = "MSTCOD" And pval.CharPressed = 9 Then
                If oMat1.Columns("MSTCOD").Cells(pval.Row).Specific.VALUE = "" Then
                    oMat1.Columns("MSTCOD").Cells(pval.Row).CLICK ct_Regular
                    Sbo_Application.ActivateMenuItem ("7425")
                    BubbleEvent = False
                Else
                    If MDC_SetMod.Value_ChkYn("[@PH_PY001A]", "Code", "'" + oMat1.Columns("MSTCOD").Cells(pval.Row).Specific.VALUE + "'") = True Then
                        oMat1.Columns("MSTCOD").Cells(pval.Row).CLICK ct_Regular
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                End If
            End If
'----------------------------------------------------------
        Case et_GOT_FOCUS: '//3
            Select Case pval.ItemUID
            Case "Mat1", "Mat2"
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
'----------------------------------------------------------
        Case et_CLICK: '//6
            If pval.BeforeAction = True Then
                Select Case pval.ItemUID
                Case "Mat1", "Mat2"
                    If pval.Row > 0 Then
                        Call oMat1.SelectRow(pval.Row, True, False)

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
                    If pval.ItemUID = "Mat1" And pval.ColUID = "MSTCOD" Then
                        If oMat1.Columns("MSTCOD").Cells(pval.Row).Specific.VALUE <> "" Then
                            oMat1.Columns("MSTNAM").Cells(pval.Row).Specific.VALUE = MDC_SetMod.Get_ReData("U_FULLNAME", "Code", "[@PH_PY001A]", "'" & oMat1.Columns("MSTCOD").Cells(pval.Row).Specific.VALUE & "'")
                            oMat1.Columns("DPTCOD").Cells(pval.Row).Specific.VALUE = MDC_SetMod.Get_ReData("U_TeamCode", "Code", "[@PH_PY001A]", "'" & oMat1.Columns("MSTCOD").Cells(pval.Row).Specific.VALUE & "'")
                            oMat1.Columns("DPTNAM").Cells(pval.Row).Specific.VALUE = MDC_SetMod.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" & oMat1.Columns("DPTCOD").Cells(pval.Row).Specific.VALUE & "'", " AND Code = '1'")
                        End If
                        Call PH_PY109_AddMatrixRow
                        oMat1.Columns("MSTCOD").Cells(pval.Row).CLICK ct_Regular
                    End If
                    
                    If pval.ItemUID = "Mat1" And pval.ColUID = "MSTCOD" Then
                        Call PH_PY109_AddMatrixRow
                        Call oMat1.Columns("MSTCOD").Cells(pval.Row).CLICK(ct_Regular)
                    End If
                End If
            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_MATRIX_LOAD: '//11
            If pval.BeforeAction = True Then
            ElseIf pval.BeforeAction = False Then
                oMat1.LoadFromDataSource

                Call PH_PY109_FormItemEnabled
                Call PH_PY109_AddMatrixRow
                Call PH_PY109_TitleSetting_Matrix
                
                Call Sum_Display
                
                
            End If
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
                Set oDS_PH_PY109A = Nothing
                Set oDS_PH_PY109B = Nothing
                Set oDS_PH_PY109Z = Nothing
        
                Set oMat1 = Nothing
                Set oMat2 = Nothing
                
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
            If pval.BeforeAction = True Then

            ElseIf pval.Before_Action = False Then
                If pval.ItemUID = "Mat1" Then
                    If pval.ColUID = "MSTCOD" Then      '// 프로젝트 코드
                        Call MDC_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY109B", "U_MSTCOD,U_MSTNAM,U_DPTCOD,U_DPTNAM", "Mat1", pval.Row)
                        Call PH_PY109_AddMatrixRow
                        Call oMat1.Columns("MSTCOD").Cells(pval.Row).CLICK(ct_Regular)
                        Call oMat1.AutoResizeColumns
                    End If
                End If
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
                Call AuthorityCheck(oForm, "CLTCOD", "@PH_PY109A", "Code")      '//접속자 권한에 따른 사업장 보기
        End Select
    ElseIf (pval.BeforeAction = False) Then
        Select Case pval.MenuUID
            Case "1283":
                oForm.Mode = fm_ADD_MODE
                Call PH_PY109_FormItemEnabled
                
                
'                Call PH_PY109_AddMatrixRow
            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//문서찾기
                Call PH_PY109_FormItemEnabled
                Call Sum_Display
'                Call PH_PY109_AddMatrixRow
                oForm.Items("Code").CLICK ct_Regular
            Case "1282":    '//문서추가
                Call PH_PY109_FormItemEnabled
'                Call PH_PY109_AddMatrixRow
                Call Sum_Display
            Case "1288", "1289", "1290", "1291":
                Call PH_PY109_FormItemEnabled
                
            Case "1293" '// 행삭제
                '// [MAT1 용]
                 If oMat1.RowCount <> oMat1.VisualRowCount Then
                    oMat1.FlushToDataSource

                    While (i <= oDS_PH_PY109B.Size - 1)
                        If oDS_PH_PY109B.GetValue("U_MSTCOD", i) = "" Then
                            oDS_PH_PY109B.RemoveRecord (i)
                            i = 0
                        Else
                            i = i + 1
                        End If
                    Wend

                    For i = 0 To oDS_PH_PY109B.Size
                        Call oDS_PH_PY109B.setValue("U_LineNum", i, i + 1)
                    Next i

                    oMat1.LoadFromDataSource
                End If

'                Call PH_PY109_AddMatrixRow
               
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
    
    Exit Sub
Raise_RightClickEvent_Error:

    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY109_AddMatrixRow()
    Dim oRow    As Long
    
    On Error GoTo PH_PY109_AddMatrixRow_Error
    
    Call oForm.Freeze(True)
    
'    '//[Mat1 용]
    oMat1.FlushToDataSource
    oRow = oMat1.VisualRowCount

    If oMat1.VisualRowCount > 0 Then
        If Trim(oDS_PH_PY109B.GetValue("U_MSTCOD", oRow - 1)) <> "" Then
            If oDS_PH_PY109B.Size <= oMat1.VisualRowCount Then
                oDS_PH_PY109B.InsertRecord (oRow)
            End If
            oDS_PH_PY109B.Offset = oRow
            oDS_PH_PY109B.setValue "U_LineNum", oRow, oRow + 1
            oDS_PH_PY109B.setValue "U_MSTCOD", oRow, ""
            oDS_PH_PY109B.setValue "U_MSTNAM", oRow, ""
            oDS_PH_PY109B.setValue "U_DPTCOD", oRow, ""
            oDS_PH_PY109B.setValue "U_DPTNAM", oRow, ""
            oDS_PH_PY109B.setValue "U_AMT01", oRow, 0
            oDS_PH_PY109B.setValue "U_AMT02", oRow, 0
            oDS_PH_PY109B.setValue "U_AMT03", oRow, 0
            oDS_PH_PY109B.setValue "U_AMT04", oRow, 0
            oDS_PH_PY109B.setValue "U_AMT05", oRow, 0
            oDS_PH_PY109B.setValue "U_AMT06", oRow, 0
            oDS_PH_PY109B.setValue "U_AMT07", oRow, 0
            oDS_PH_PY109B.setValue "U_AMT08", oRow, 0
            oDS_PH_PY109B.setValue "U_AMT09", oRow, 0
            oDS_PH_PY109B.setValue "U_AMT10", oRow, 0
            oDS_PH_PY109B.setValue "U_AMT11", oRow, 0
            oDS_PH_PY109B.setValue "U_AMT12", oRow, 0
            oMat1.LoadFromDataSource
        Else
            oDS_PH_PY109B.Offset = oRow - 1
            oDS_PH_PY109B.setValue "U_LineNum", oRow - 1, oRow
            oDS_PH_PY109B.setValue "U_MSTCOD", oRow - 1, ""
            oDS_PH_PY109B.setValue "U_MSTNAM", oRow - 1, ""
            oDS_PH_PY109B.setValue "U_DPTCOD", oRow - 1, ""
            oDS_PH_PY109B.setValue "U_DPTNAM", oRow - 1, ""
            oDS_PH_PY109B.setValue "U_AMT01", oRow - 1, 0
            oDS_PH_PY109B.setValue "U_AMT02", oRow - 1, 0
            oDS_PH_PY109B.setValue "U_AMT03", oRow - 1, 0
            oDS_PH_PY109B.setValue "U_AMT04", oRow - 1, 0
            oDS_PH_PY109B.setValue "U_AMT05", oRow - 1, 0
            oDS_PH_PY109B.setValue "U_AMT06", oRow - 1, 0
            oDS_PH_PY109B.setValue "U_AMT07", oRow - 1, 0
            oDS_PH_PY109B.setValue "U_AMT08", oRow - 1, 0
            oDS_PH_PY109B.setValue "U_AMT09", oRow - 1, 0
            oDS_PH_PY109B.setValue "U_AMT10", oRow - 1, 0
            oDS_PH_PY109B.setValue "U_AMT11", oRow - 1, 0
            oDS_PH_PY109B.setValue "U_AMT12", oRow - 1, 0
            oMat1.LoadFromDataSource
        End If
    ElseIf oMat1.VisualRowCount = 0 Then
        oDS_PH_PY109B.Offset = oRow
        oDS_PH_PY109B.setValue "U_LineNum", oRow, oRow + 1
        oDS_PH_PY109B.setValue "U_MSTCOD", oRow, ""
        oDS_PH_PY109B.setValue "U_MSTNAM", oRow, ""
        oDS_PH_PY109B.setValue "U_DPTCOD", oRow, ""
        oDS_PH_PY109B.setValue "U_DPTNAM", oRow, ""
        oDS_PH_PY109B.setValue "U_AMT01", oRow, 0
        oDS_PH_PY109B.setValue "U_AMT02", oRow, 0
        oDS_PH_PY109B.setValue "U_AMT03", oRow, 0
        oDS_PH_PY109B.setValue "U_AMT04", oRow, 0
        oDS_PH_PY109B.setValue "U_AMT05", oRow, 0
        oDS_PH_PY109B.setValue "U_AMT06", oRow, 0
        oDS_PH_PY109B.setValue "U_AMT07", oRow, 0
        oDS_PH_PY109B.setValue "U_AMT08", oRow, 0
        oDS_PH_PY109B.setValue "U_AMT09", oRow, 0
        oDS_PH_PY109B.setValue "U_AMT10", oRow, 0
        oDS_PH_PY109B.setValue "U_AMT11", oRow, 0
        oDS_PH_PY109B.setValue "U_AMT12", oRow, 0
        oMat1.LoadFromDataSource
    End If
    
    Call oForm.Freeze(False)
    Exit Sub
PH_PY109_AddMatrixRow_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY109_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Sub PH_PY109_FormClear()
On Error GoTo PH_PY109_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY109'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PH_PY109_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY109_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY109_DataValidCheck(ItemUID As String) As Boolean
On Error GoTo PH_PY109_DataValidCheck_Error
    PH_PY109_DataValidCheck = False
    Dim i       As Long
    Dim sQry    As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '//----------------------------------------------------------------------------------
    '//필수 체크
    '//----------------------------------------------------------------------------------
    If Trim(oDS_PH_PY109A.GetValue("U_CLTCOD", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "사업장은 필수입니다.", bmt_Short, True
        oForm.Items("CLTCOD").CLICK ct_Regular
        Exit Function
    End If
    
    If Trim(oDS_PH_PY109A.GetValue("U_YM", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "귀속년월은 필수입니다.", bmt_Short, True
        oForm.Items("YM").CLICK ct_Regular
        Exit Function
    End If
    
    If Trim(oDS_PH_PY109A.GetValue("U_JOBTYP", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "지급종류는 필수입니다.", bmt_Short, True
        oForm.Items("JOBTYP").CLICK ct_Regular
        Exit Function
    End If
    
    If Trim(oDS_PH_PY109A.GetValue("U_JOBGBN", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "지급구분은 필수입니다.", bmt_Short, True
        oForm.Items("JOBGBN").CLICK ct_Regular
        Exit Function
    End If
    
    If Trim(oDS_PH_PY109A.GetValue("U_JOBTRG", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "지급대상은 필수입니다.", bmt_Short, True
        oForm.Items("JOBTRG").CLICK ct_Regular
        Exit Function
    End If
    
    
    Call oDS_PH_PY109A.setValue("Code", 0, Trim$(oDS_PH_PY109A.GetValue("U_CLTCOD", 0)) & Right(Trim$(oDS_PH_PY109A.GetValue("U_YM", 0)), 4) _
        & Trim$(oDS_PH_PY109A.GetValue("U_JOBTYP", 0)) & Trim$(oDS_PH_PY109A.GetValue("U_JOBGBN", 0)) & Trim$(oDS_PH_PY109A.GetValue("U_JOBTRG", 0)))
    Call oDS_PH_PY109A.setValue("Name", 0, Trim$(oDS_PH_PY109A.GetValue("COde", 0)))
    
    If ItemUID = "1" And oForm.Mode = fm_ADD_MODE Then
        If MDC_SetMod.Get_ReData("Code", "Code", "[@PH_PY109A]", "'" & Trim$(oDS_PH_PY109A.GetValue("COde", 0)) & "'") <> "" Then
            Sbo_Application.SetStatusBarMessage "이미 저장되어져 있는 헤더의 내용과 일치합니다", bmt_Short, True
            Exit Function
        End If
    ElseIf ItemUID = "Btn_Set" Then
        If MDC_SetMod.Get_ReData("Code", "Code", "[@PH_PY109A]", "'" & Trim$(oDS_PH_PY109A.GetValue("COde", 0)) & "'") <> "" Then
            Sbo_Application.SetStatusBarMessage "이미 저장되어져 있는 헤더의 내용과 일치합니다", bmt_Short, True
            Exit Function
        End If
    End If
        
    oMat1.FlushToDataSource
    
    
    If ItemUID = "1" Then
        '// Matrix 마지막 행 삭제(DB 저장시)
        If oDS_PH_PY109B.Size > 1 Then
            oDS_PH_PY109B.RemoveRecord (oDS_PH_PY109B.Size - 1)
        Else
            Sbo_Application.SetStatusBarMessage "라인이 존재하지 않습니다.", bmt_Short, True
            PH_PY109_DataValidCheck = False
            Exit Function
        End If
    End If
    oMat1.LoadFromDataSource
    
    PH_PY109_DataValidCheck = True
    
    Exit Function
    
    Set oRecordSet = Nothing
    
    
PH_PY109_DataValidCheck_Error:
    Set oRecordSet = Nothing
    PH_PY109_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY109_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Function PH_PY109_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY109_Validate_Error
    PH_PY109_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY109A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", bmt_Short, True
        PH_PY109_Validate = False
        GoTo PH_PY109_Validate_Exit
    End If
'
    If ValidateType = "수정" Then

    ElseIf ValidateType = "행삭제" Then

    ElseIf ValidateType = "취소" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY109_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY109_Validate_Error:
    PH_PY109_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY109_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Function PH_PY109_LoadData_SudangGongje()
    Dim i           As Long
    Dim sQry        As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    Dim oCombo      As SAPbouiCOM.ComboBox

    On Error GoTo PH_PY109_DataLoad_ERROR
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    oForm.Freeze (True)
      
    If oDS_PH_PY109Z.Size > 0 Then
        oDS_PH_PY109Z.Clear
    End If

    sCode = Trim(oDS_PH_PY109A.GetValue("Code", 0))
    sCLTCOD = Trim(oDS_PH_PY109A.GetValue("U_CLTCOD", 0))
    sYM = Trim(oDS_PH_PY109A.GetValue("U_YM", 0))
    sJOBTYP = Trim(oDS_PH_PY109A.GetValue("U_JOBTYP", 0))
    sJOBGBN = Trim(oDS_PH_PY109A.GetValue("U_JOBGBN", 0))
    sJOBTRG = Trim(oDS_PH_PY109A.GetValue("U_JOBTRG", 0))
        
    '// 수당, 공제 테이블 고정:V, 상여:Y 인 값을 임시테이블에 넣는다
    sQry = "EXEC PH_PY109 '" & Trim(oDS_PH_PY109A.GetValue("U_CLTCOD", 0)) & "' , '" & Trim(oDS_PH_PY109A.GetValue("U_YM", 0)) & "' , '"
    sQry = sQry & Trim(oDS_PH_PY109A.GetValue("U_JOBTYP", 0)) & "' , '" & Trim(oDS_PH_PY109A.GetValue("U_JOBGBN", 0)) & "' , '"
    sQry = sQry & Trim(oDS_PH_PY109A.GetValue("U_JOBTRG", 0)) & "'"
    
    oRecordSet.DoQuery sQry

    oForm.Mode = fm_FIND_MODE
    
    Set oCombo = oForm.Items("CLTCOD").Specific
    oCombo.Select "" & sCLTCOD & "", psk_ByValue
    oForm.Items("YM").Specific.VALUE = sYM
    Set oCombo = oForm.Items("JOBTYP").Specific
    oCombo.Select "" & sJOBTYP & "", psk_ByValue
    Set oCombo = oForm.Items("JOBGBN").Specific
    oCombo.Select "" & sJOBGBN & "", psk_ByValue
    Set oCombo = oForm.Items("JOBTRG").Specific
    oCombo.Select "" & sJOBTRG & "", psk_ByValue
    
    oForm.Items("1").CLICK ct_Regular
    
    oForm.Freeze (False)
    Exit Function
    
PH_PY109_DataLoad_ERROR:
    oForm.Freeze (False)
    Sbo_Application.SetStatusBarMessage "PH_PY109_DataLoad_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Function PH_PY109_LoadData_User()
    Dim i           As Long
    Dim sQry        As String
    Dim oRecordSet  As SAPbobsCOM.Recordset

    On Error GoTo PH_PY109_DataLoad_ERROR

    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    oForm.Freeze (True)
'
'    sQry = ""
'    oDS_PH_PY109_Grid1.Rows.Clear
'
'    If oDS_PH_PY109B.Size > 1 Then
'        oDS_PH_PY109_Grid1.Rows.Add (oDS_PH_PY109B.Size)
'        For i = 0 To oDS_PH_PY109B.Size - 1
'            oDS_PH_PY109_Grid1.setValue "U_MSTCOD", i, oDS_PH_PY109B.GetValue("U_MSTCOD", i)
'            oDS_PH_PY109_Grid1.setValue "U_MSTNAM", i, oDS_PH_PY109B.GetValue("U_MSTNAM", i)
'            oDS_PH_PY109_Grid1.setValue "U_DPTCOD", i, oDS_PH_PY109B.GetValue("U_DPTCOD", i)
'            oDS_PH_PY109_Grid1.setValue "U_DPTNAM", i, oDS_PH_PY109B.GetValue("U_DPTNAM", i)
'            oDS_PH_PY109_Grid1.setValue "U_AMT01", i, oDS_PH_PY109B.GetValue("U_AMT01", i)
'            oDS_PH_PY109_Grid1.setValue "U_AMT02", i, oDS_PH_PY109B.GetValue("U_AMT02", i)
'            oDS_PH_PY109_Grid1.setValue "U_AMT03", i, oDS_PH_PY109B.GetValue("U_AMT03", i)
'            oDS_PH_PY109_Grid1.setValue "U_AMT04", i, oDS_PH_PY109B.GetValue("U_AMT04", i)
'            oDS_PH_PY109_Grid1.setValue "U_AMT05", i, oDS_PH_PY109B.GetValue("U_AMT05", i)
'            oDS_PH_PY109_Grid1.setValue "U_AMT06", i, oDS_PH_PY109B.GetValue("U_AMT06", i)
'            oDS_PH_PY109_Grid1.setValue "U_AMT07", i, oDS_PH_PY109B.GetValue("U_AMT07", i)
'            oDS_PH_PY109_Grid1.setValue "U_AMT08", i, oDS_PH_PY109B.GetValue("U_AMT08", i)
'            oDS_PH_PY109_Grid1.setValue "U_AMT09", i, oDS_PH_PY109B.GetValue("U_AMT09", i)
'            oDS_PH_PY109_Grid1.setValue "U_AMT10", i, oDS_PH_PY109B.GetValue("U_AMT10", i)
'            oDS_PH_PY109_Grid1.setValue "U_AMT11", i, oDS_PH_PY109B.GetValue("U_AMT11", i)
'            oDS_PH_PY109_Grid1.setValue "U_AMT12", i, oDS_PH_PY109B.GetValue("U_AMT12", i)
'        Next
'    End If

'    oDS_PH_PY109_Grid1.Rows.Add (1)

'    oGrid1.AutoResizeColumns

    oForm.Freeze (False)
    Exit Function

PH_PY109_DataLoad_ERROR:
    oForm.Freeze (False)
    Sbo_Application.SetStatusBarMessage "PH_PY109_LoadData_User_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY109_TitleSetting_Matrix()
    Dim i               As Long
    Dim iCount          As Long
    
    On Error GoTo Error_Message
    
    oForm.Freeze True
    iCount = 0
    For i = 5 To oMat1.Columns.Count - 1
        If oDS_PH_PY109Z.Size >= 1 Then
            If Trim(oDS_PH_PY109Z.GetValue("U_PDName", 0)) = "" Then
                oMat1.Columns(1).Editable = False
                oMat1.Columns(i).TitleObject.Caption = ""
                oMat1.Columns(i).Editable = False
            Else
                oMat1.Columns(1).Editable = True
                If i > 4 And i <= oDS_PH_PY109Z.Size + 4 Then
                    oMat1.Columns(i).TitleObject.Caption = Trim(oDS_PH_PY109Z.GetValue("U_PDName", iCount))
                    oMat1.Columns(i).Editable = True
                    iCount = iCount + 1
                Else
                    oMat1.Columns(i).TitleObject.Caption = ""
                    oMat1.Columns(i).Editable = False
                End If
            End If
        Else
            oMat1.Columns(1).Editable = False
            oMat1.Columns(i).TitleObject.Caption = ""
            oMat1.Columns(i).Editable = False
        End If
    Next i
    

    oMat1.AutoResizeColumns

    oForm.Freeze False


    Exit Sub

Error_Message:
    oForm.Freeze False
    Sbo_Application.SetStatusBarMessage "PH_PY109_TitleSetting_Matrix Error : " & Space(10) & Err.Description, bmt_Short, True
End Sub



Private Sub PH_PY109_Excel_Upload()
    
    Dim i               As Long
    Dim j               As Long
    Dim k               As Long
    Dim CheckLine       As Long
    Dim CheckYN         As Boolean
    Dim sPrice          As String
    Dim sFile           As String
    Dim OneRec          As String
    Dim sQry            As String
    Dim ErrNum          As Integer
    Dim xl              As Excel.Application
    Dim xlwb            As Excel.Workbook
    Dim xlsh            As Excel.Worksheet
    
    Dim oEdit           As SAPbouiCOM.EditText
    Dim oForm           As SAPbouiCOM.Form
    
    Dim V_StatusCnt&, oProValue&, TOTCNT&, tRow&  '//progbar
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Dim Amt01 As Long
    Dim Amt02 As Long
    Dim Amt03 As Long
    Dim Amt04 As Long
    Dim Amt05 As Long
    Dim Amt06 As Long
    Dim Amt07 As Long
    Dim Amt08 As Long
    Dim Amt09 As Long
    Dim Amt10 As Long
    Dim Amt11 As Long
    Dim Amt12 As Long
    
    On Error GoTo Err
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    Set oForm = Sbo_Application.Forms.ActiveForm
    
    If oMat1.Columns("MSTCOD").Editable = False Then
        Sbo_Application.StatusBar.SetText "항목 설정이 되지 않았습니다.", bmt_Short, smt_Error
        Exit Sub
    End If
    
    sFile = FileListBoxForm.OpenDialog(FileListBoxForm, "*.xls", "파일선택", "C:\")
    
    If sFile = "" Then
        Exit Sub
    Else
        If Right(Replace(sFile, Chr$(0), ""), 3) <> "xls" And Right(Replace(sFile, Chr$(0), ""), 4) <> "xlsx" Then
            Sbo_Application.StatusBar.SetText "엑셀파일이 아닙니다.", bmt_Short, smt_Error
            Exit Sub
        End If
    End If
    
    If Not oProgBar Is Nothing Then
        oProgBar.Stop
        Set oProgBar = Nothing
    End If
    
    '엑셀 Object 연결
    Set xl = CreateObject("excel.application")
    Set xlwb = xl.Workbooks.Open(sFile, , True)
    Set xlsh = xlwb.Worksheets("급상여변동")
    
    
    If xlsh.UsedRange.Columns.Count <= 2 Then
        Sbo_Application.StatusBar.SetText "항목이 없습니다.", bmt_Short, smt_Error
        xlwb.Close
        Set xlwb = Nothing: Set xl = Nothing: Set xlsh = Nothing
        Exit Sub
    End If

    If xlsh.Cells(1, 1).VALUE <> "사번" Then
        Sbo_Application.StatusBar.SetText "A열 첫번째 행 타이틀은 사번", bmt_Short, smt_Error
        xlwb.Close
        Set xlwb = Nothing: Set xl = Nothing: Set xlsh = Nothing
        Exit Sub
    End If

    If xlsh.Cells(1, 2).VALUE <> "성명" Then
        Sbo_Application.StatusBar.SetText "B열 두번째 행 타이틀은 성명", bmt_Short, smt_Error
        xlwb.Close
        Set xlwb = Nothing: Set xl = Nothing: Set xlsh = Nothing
        Exit Sub
    End If
    
    '프로그레스 바    ///////////////////////////////////////
    Set oProgBar = Sbo_Application.StatusBar.CreateProgressBar("데이터 읽는중...!", 50, False)
    
    '최대값 구하기
    TOTCNT = xlsh.UsedRange.Rows.Count - 1
   
    V_StatusCnt = Round(TOTCNT / 50, 0)
    oProValue = 1
    tRow = 1
    '/////////////////////////////////////////////////////
    
    For i = 2 To xlsh.UsedRange.Rows.Count
        '//사번 존재 여부 체크
        If Value_ChkYn("[@PH_PY001A]", "Code", "'" & Trim(xlsh.Cells(i, 1)) & "'") = True Then
            ErrNum = 1
            GoTo Err
        Else
            CheckYN = False
            
            For j = 0 To oDS_PH_PY109B.Size - 1
                If Trim(xlsh.Cells(i, 1)) = Trim(oDS_PH_PY109B.GetValue("U_MSTCOD", j)) Then
                    CheckYN = True
                    CheckLine = j
                    Exit For
                End If
            Next j
            
            '//마지막행 제거
            If Trim(oDS_PH_PY109B.GetValue("U_MSTCOD", oDS_PH_PY109B.Size - 1)) = "" Then
                oDS_PH_PY109B.RemoveRecord (oDS_PH_PY109B.Size - 1)
            End If

            '// 사원마스터에서 사번에 대한 정보 가져오기
            sQry = "select U_FullName, U_TeamCode, U_CodeNm "
            sQry = sQry & " FROM [@PH_PY001A] T0 INNER JOIN [@Ps_HR200L] T1 ON T0.U_Teamcode = T1.U_Code"
            sQry = sQry & " WHERE T0.Code='" & Trim(xlsh.Cells(i, 1)) & "'"
            oRecordSet.DoQuery sQry
            
            '//Mat1에 업로드엑셀 사번이 존재 유무 CheckYN
            If CheckYN = True Then
                oDS_PH_PY109B.setValue "U_MSTNAM", CheckLine, Trim$(oRecordSet.Fields(0).VALUE)
                oDS_PH_PY109B.setValue "U_DPTCOD", CheckLine, Trim$(oRecordSet.Fields(1).VALUE)
                oDS_PH_PY109B.setValue "U_DPTNAM", CheckLine, Trim$(oRecordSet.Fields(2).VALUE)
                
                Amt01 = xlsh.Cells(i, ValueCheck(oMat1, xlsh, "AMT01"))
                If Amt01 <> 0 Then
                    oDS_PH_PY109B.setValue "U_AMT01", CheckLine, Amt01
                End If

                Amt02 = xlsh.Cells(i, ValueCheck(oMat1, xlsh, "AMT02"))
                If Amt02 <> 0 Then
                    oDS_PH_PY109B.setValue "U_AMT02", CheckLine, Amt02
                End If

                Amt03 = xlsh.Cells(i, ValueCheck(oMat1, xlsh, "AMT03"))
                If Amt03 <> 0 Then
                    oDS_PH_PY109B.setValue "U_AMT03", CheckLine, Amt03
                End If

                Amt04 = xlsh.Cells(i, ValueCheck(oMat1, xlsh, "AMT04"))
                If Amt04 <> 0 Then
                    oDS_PH_PY109B.setValue "U_AMT04", CheckLine, Amt04
                End If

                Amt05 = xlsh.Cells(i, ValueCheck(oMat1, xlsh, "AMT05"))
                If Amt05 <> 0 Then
                    oDS_PH_PY109B.setValue "U_AMT05", CheckLine, Amt05
                End If

                Amt06 = xlsh.Cells(i, ValueCheck(oMat1, xlsh, "AMT06"))
                If Amt06 <> 0 Then
                    oDS_PH_PY109B.setValue "U_AMT06", CheckLine, Amt06
                End If

                Amt07 = xlsh.Cells(i, ValueCheck(oMat1, xlsh, "AMT07"))
                If Amt07 <> 0 Then
                    oDS_PH_PY109B.setValue "U_AMT07", CheckLine, Amt07
                End If

                Amt08 = xlsh.Cells(i, ValueCheck(oMat1, xlsh, "AMT08"))
                If Amt08 <> 0 Then
                    oDS_PH_PY109B.setValue "U_AMT08", CheckLine, Amt08
                End If

                Amt09 = xlsh.Cells(i, ValueCheck(oMat1, xlsh, "AMT09"))
                If Amt09 <> 0 Then
                    oDS_PH_PY109B.setValue "U_AMT09", CheckLine, Amt09
                End If

                Amt10 = xlsh.Cells(i, ValueCheck(oMat1, xlsh, "AMT10"))
                If Amt10 <> 0 Then
                    oDS_PH_PY109B.setValue "U_AMT10", CheckLine, Amt10
                End If

                Amt11 = xlsh.Cells(i, ValueCheck(oMat1, xlsh, "AMT11"))
                If Amt11 <> 0 Then
                    oDS_PH_PY109B.setValue "U_AMT11", CheckLine, Amt11
                End If

                Amt12 = xlsh.Cells(i, ValueCheck(oMat1, xlsh, "AMT12"))
                If Amt12 <> 0 Then
                    oDS_PH_PY109B.setValue "U_AMT12", CheckLine, Amt12
                End If
                
'                If TitleCheck(oMat1, xlsh, "AMT01") = True Then
'                    oDS_PH_PY109B.setValue "U_AMT01", CheckLine, xlsh.Cells(i, ValueCheck(oMat1, xlsh, "AMT01"))
'                End If
'                If TitleCheck(oMat1, xlsh, "AMT02") = True Then
'                    oDS_PH_PY109B.setValue "U_AMT02", CheckLine, xlsh.Cells(i, ValueCheck(oMat1, xlsh, "AMT02"))
'                End If
'                If TitleCheck(oMat1, xlsh, "AMT03") = True Then
'                    oDS_PH_PY109B.setValue "U_AMT03", CheckLine, xlsh.Cells(i, ValueCheck(oMat1, xlsh, "AMT03"))
'                End If
'                If TitleCheck(oMat1, xlsh, "AMT04") = True Then
'                    oDS_PH_PY109B.setValue "U_AMT04", CheckLine, xlsh.Cells(i, ValueCheck(oMat1, xlsh, "AMT04"))
'                End If
'                If TitleCheck(oMat1, xlsh, "AMT05") = True Then
'                    oDS_PH_PY109B.setValue "U_AMT05", CheckLine, xlsh.Cells(i, ValueCheck(oMat1, xlsh, "AMT05"))
'                End If
'                If TitleCheck(oMat1, xlsh, "AMT06") = True Then
'                    oDS_PH_PY109B.setValue "U_AMT06", CheckLine, xlsh.Cells(i, ValueCheck(oMat1, xlsh, "AMT06"))
'                End If
'                If TitleCheck(oMat1, xlsh, "AMT07") = True Then
'                    oDS_PH_PY109B.setValue "U_AMT07", CheckLine, xlsh.Cells(i, ValueCheck(oMat1, xlsh, "AMT07"))
'                End If
'                If TitleCheck(oMat1, xlsh, "AMT08") = True Then
'                    oDS_PH_PY109B.setValue "U_AMT08", CheckLine, xlsh.Cells(i, ValueCheck(oMat1, xlsh, "AMT08"))
'                End If
'                If TitleCheck(oMat1, xlsh, "AMT09") = True Then
'                    oDS_PH_PY109B.setValue "U_AMT09", CheckLine, xlsh.Cells(i, ValueCheck(oMat1, xlsh, "AMT09"))
'                End If
'                If TitleCheck(oMat1, xlsh, "AMT10") = True Then
'                    oDS_PH_PY109B.setValue "U_AMT10", CheckLine, xlsh.Cells(i, ValueCheck(oMat1, xlsh, "AMT10"))
'                End If
'                If TitleCheck(oMat1, xlsh, "AMT11") = True Then
'                    oDS_PH_PY109B.setValue "U_AMT11", CheckLine, xlsh.Cells(i, ValueCheck(oMat1, xlsh, "AMT11"))
'                End If
'                If TitleCheck(oMat1, xlsh, "AMT12") = True Then
'                    oDS_PH_PY109B.setValue "U_AMT12", CheckLine, xlsh.Cells(i, ValueCheck(oMat1, xlsh, "AMT12"))
'                End If
            Else        '//새로 사번 추가
'                If oDS_PH_PY109B.Size <= oMat1.VisualRowCount Then
                    oDS_PH_PY109B.InsertRecord (oDS_PH_PY109B.Size)
'                End If
                oDS_PH_PY109B.Offset = oDS_PH_PY109B.Size - 1
                oDS_PH_PY109B.setValue "U_MSTCOD", oDS_PH_PY109B.Size - 1, xlsh.Cells(i, 1)
                oDS_PH_PY109B.setValue "U_MSTNAM", oDS_PH_PY109B.Size - 1, Trim$(oRecordSet.Fields(0).VALUE)
                oDS_PH_PY109B.setValue "U_DPTCOD", oDS_PH_PY109B.Size - 1, Trim$(oRecordSet.Fields(1).VALUE)
                oDS_PH_PY109B.setValue "U_DPTNAM", oDS_PH_PY109B.Size - 1, Trim$(oRecordSet.Fields(2).VALUE)
                If TitleCheck(oMat1, xlsh, "AMT01") = True Then
                    oDS_PH_PY109B.setValue "U_AMT01", oDS_PH_PY109B.Size - 1, xlsh.Cells(i, ValueCheck(oMat1, xlsh, "AMT01"))
                End If
                If TitleCheck(oMat1, xlsh, "AMT02") = True Then
                    oDS_PH_PY109B.setValue "U_AMT02", oDS_PH_PY109B.Size - 1, xlsh.Cells(i, ValueCheck(oMat1, xlsh, "AMT02"))
                End If
                If TitleCheck(oMat1, xlsh, "AMT03") = True Then
                    oDS_PH_PY109B.setValue "U_AMT03", oDS_PH_PY109B.Size - 1, xlsh.Cells(i, ValueCheck(oMat1, xlsh, "AMT03"))
                End If
                If TitleCheck(oMat1, xlsh, "AMT04") = True Then
                    oDS_PH_PY109B.setValue "U_AMT04", oDS_PH_PY109B.Size - 1, xlsh.Cells(i, ValueCheck(oMat1, xlsh, "AMT04"))
                End If
                If TitleCheck(oMat1, xlsh, "AMT05") = True Then
                    oDS_PH_PY109B.setValue "U_AMT05", oDS_PH_PY109B.Size - 1, xlsh.Cells(i, ValueCheck(oMat1, xlsh, "AMT05"))
                End If
                If TitleCheck(oMat1, xlsh, "AMT06") = True Then
                    oDS_PH_PY109B.setValue "U_AMT06", oDS_PH_PY109B.Size - 1, xlsh.Cells(i, ValueCheck(oMat1, xlsh, "AMT06"))
                End If
                If TitleCheck(oMat1, xlsh, "AMT07") = True Then
                    oDS_PH_PY109B.setValue "U_AMT07", oDS_PH_PY109B.Size - 1, xlsh.Cells(i, ValueCheck(oMat1, xlsh, "AMT07"))
                End If
                If TitleCheck(oMat1, xlsh, "AMT08") = True Then
                    oDS_PH_PY109B.setValue "U_AMT08", oDS_PH_PY109B.Size - 1, xlsh.Cells(i, ValueCheck(oMat1, xlsh, "AMT08"))
                End If
                If TitleCheck(oMat1, xlsh, "AMT09") = True Then
                    oDS_PH_PY109B.setValue "U_AMT09", oDS_PH_PY109B.Size - 1, xlsh.Cells(i, ValueCheck(oMat1, xlsh, "AMT09"))
                End If
                If TitleCheck(oMat1, xlsh, "AMT10") = True Then
                    oDS_PH_PY109B.setValue "U_AMT10", oDS_PH_PY109B.Size - 1, xlsh.Cells(i, ValueCheck(oMat1, xlsh, "AMT10"))
                End If
                If TitleCheck(oMat1, xlsh, "AMT11") = True Then
                    oDS_PH_PY109B.setValue "U_AMT11", oDS_PH_PY109B.Size - 1, xlsh.Cells(i, ValueCheck(oMat1, xlsh, "AMT11"))
                End If
                If TitleCheck(oMat1, xlsh, "AMT12") = True Then
                    oDS_PH_PY109B.setValue "U_AMT12", oDS_PH_PY109B.Size - 1, xlsh.Cells(i, ValueCheck(oMat1, xlsh, "AMT12"))
                End If
            End If

            If (TOTCNT > 50 And tRow = oProValue * V_StatusCnt) Or TOTCNT <= 50 Then
                oProgBar.Text = tRow & "/ " & TOTCNT & " 건 처리중...!"
                oProValue = oProValue + 1
                oProgBar.VALUE = oProValue
            End If
            tRow = tRow + 1
        End If
    Next i

    '//수당, 공제 항목 이외의 값은 0으로 처리
    For i = 0 To oDS_PH_PY109B.Size - 1
        For j = 9 + oDS_PH_PY109Z.Size To oDS_PH_PY109B.Fields.Count - 1
            oDS_PH_PY109B.setValue j, i, 0
        Next j
        Call oDS_PH_PY109B.setValue("U_LineNum", i, i + 1)                  '//라인번호 재정의
    Next i
    
    oMat1.LoadFromDataSource
    oMat1.AutoResizeColumns
    
    
    oProgBar.Stop
    Sbo_Application.StatusBar.SetText "엑셀을 불러왔습니다.", bmt_Short, smt_Success
    
    '액셀개체 닫음
    xlwb.Close
    Set xlwb = Nothing: Set xl = Nothing: Set xlsh = Nothing
    Set oProgBar = Nothing
    '진행바 초기화
    oForm.Mode = fm_UPDATE_MODE
    oForm.Items("CLTCOD").Enabled = False
    oForm.Items("YM").Enabled = False
    oForm.Items("JOBTYP").Enabled = False
    oForm.Items("JOBGBN").Enabled = False
    oForm.Items("JOBTRG").Enabled = False
    
    Exit Sub
    
Err:
    If Not oProgBar Is Nothing Then
        oProgBar.Stop
        Set oProgBar = Nothing
    End If
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText i & "행의 [" & xlsh.Cells(i, 1) & " ] 사번이 존재하지 않습니다.", bmt_Long, smt_Error
    Else
        Call Sbo_Application.StatusBar.SetText(Err.Description)
    End If
    Set oProgBar = Nothing
    xlwb.Close
    Set xlwb = Nothing: Set xl = Nothing: Set xlsh = Nothing
    
End Sub

Private Function ValueCheck(oMat1 As SAPbouiCOM.Matrix, xlsh As Excel.Worksheet, Field As String) As Long
    Dim i       As Long
    Dim check   As Boolean
    
    check = False
    For i = 3 To xlsh.UsedRange.Columns.Count
        If xlsh.Cells(1, i).VALUE <> "" Then
            If check = False Then
                If xlsh.Cells(1, i).VALUE = oMat1.Columns(Field).Title Then
                    ValueCheck = i
                    check = True
                    If ValueCheck <= 2 Then
                        ValueCheck = 50 '16384
                    End If
                    Exit For
                End If
            Else
                Sbo_Application.StatusBar.SetText "항목에 중복된 이름이 있습니다.", bmt_Long, smt_Error
            End If
        End If
    Next i
    
    '//해당 항목이 없으면 마지막 컬럼값(빈값)의 0 값을 가져오게 설정
    If ValueCheck <= 2 Then
        ValueCheck = 50 '16384
    End If
End Function


Private Function TitleCheck(oMat1 As SAPbouiCOM.Matrix, xlsh As Excel.Worksheet, Field As String) As Boolean
    Dim i       As Long
    Dim check   As Boolean
    
    check = False
    For i = 3 To xlsh.UsedRange.Columns.Count
        If xlsh.Cells(1, i).VALUE <> "" Then
            If check = False Then
                If xlsh.Cells(1, i).VALUE = oMat1.Columns(Field).Title Then
                    TitleCheck = True
                    check = True
                End If
            Else
                'Sbo_Application.StatusBar.SetText "항목에 중복된 이름이 있습니다.", bmt_Long, smt_Error
            End If
        End If
    Next i
End Function


Private Function Sum_Display() As Long
    Dim i       As Long
    
    
    Dim Amt01 As Long
    Dim Amt02 As Long
    Dim Amt03 As Long
    Dim Amt04 As Long
    Dim Amt05 As Long
    Dim Amt06 As Long
    Dim Amt07 As Long
    Dim Amt08 As Long
    Dim Amt09 As Long
    Dim Amt10 As Long
    Dim Amt11 As Long
    Dim Amt12 As Long
    
    Amt01 = 0
    Amt02 = 0
    Amt03 = 0
    Amt04 = 0
    Amt05 = 0
    Amt06 = 0
    Amt07 = 0
    Amt08 = 0
    Amt09 = 0
    Amt10 = 0
    Amt11 = 0
    Amt12 = 0
    
    oMat1.FlushToDataSource
    oMat1.LoadFromDataSource
    For i = 0 To oMat1.VisualRowCount - 1
        Amt01 = Amt01 + oDS_PH_PY109B.GetValue("U_AMT01", i)
        Amt02 = Amt02 + oDS_PH_PY109B.GetValue("U_AMT02", i)
        Amt03 = Amt03 + oDS_PH_PY109B.GetValue("U_AMT03", i)
        Amt04 = Amt04 + oDS_PH_PY109B.GetValue("U_AMT04", i)
        Amt05 = Amt05 + oDS_PH_PY109B.GetValue("U_AMT05", i)
        Amt06 = Amt06 + oDS_PH_PY109B.GetValue("U_AMT06", i)
        Amt07 = Amt07 + oDS_PH_PY109B.GetValue("U_AMT07", i)
        Amt08 = Amt08 + oDS_PH_PY109B.GetValue("U_AMT08", i)
        Amt09 = Amt09 + oDS_PH_PY109B.GetValue("U_AMT09", i)
        Amt10 = Amt10 + oDS_PH_PY109B.GetValue("U_AMT10", i)
        Amt11 = Amt11 + oDS_PH_PY109B.GetValue("U_AMT11", i)
        Amt12 = Amt12 + oDS_PH_PY109B.GetValue("U_AMT12", i)
        
    Next i
    
    oForm.Items("T1").Specific.Caption = oMat1.Columns("AMT01").TitleObject.Caption
    oForm.Items("T2").Specific.Caption = oMat1.Columns("AMT02").TitleObject.Caption
    oForm.Items("T3").Specific.Caption = oMat1.Columns("AMT03").TitleObject.Caption
    oForm.Items("T4").Specific.Caption = oMat1.Columns("AMT04").TitleObject.Caption
    oForm.Items("T5").Specific.Caption = oMat1.Columns("AMT05").TitleObject.Caption
    oForm.Items("T6").Specific.Caption = oMat1.Columns("AMT06").TitleObject.Caption
    oForm.Items("T7").Specific.Caption = oMat1.Columns("AMT07").TitleObject.Caption
    oForm.Items("T8").Specific.Caption = oMat1.Columns("AMT08").TitleObject.Caption
    oForm.Items("T9").Specific.Caption = oMat1.Columns("AMT09").TitleObject.Caption
    oForm.Items("T10").Specific.Caption = oMat1.Columns("AMT10").TitleObject.Caption
    oForm.Items("T11").Specific.Caption = oMat1.Columns("AMT11").TitleObject.Caption
    oForm.Items("T12").Specific.Caption = oMat1.Columns("AMT12").TitleObject.Caption
    
    oForm.DataSources.UserDataSources.Item("S_Amt01").VALUE = Amt01
    oForm.DataSources.UserDataSources.Item("S_Amt02").VALUE = Amt02
    oForm.DataSources.UserDataSources.Item("S_Amt03").VALUE = Amt03
    oForm.DataSources.UserDataSources.Item("S_Amt04").VALUE = Amt04
    oForm.DataSources.UserDataSources.Item("S_Amt05").VALUE = Amt05
    oForm.DataSources.UserDataSources.Item("S_Amt06").VALUE = Amt06
    oForm.DataSources.UserDataSources.Item("S_Amt07").VALUE = Amt07
    oForm.DataSources.UserDataSources.Item("S_Amt08").VALUE = Amt08
    oForm.DataSources.UserDataSources.Item("S_Amt09").VALUE = Amt09
    oForm.DataSources.UserDataSources.Item("S_Amt10").VALUE = Amt10
    oForm.DataSources.UserDataSources.Item("S_Amt11").VALUE = Amt11
    oForm.DataSources.UserDataSources.Item("S_Amt12").VALUE = Amt12
    

End Function

