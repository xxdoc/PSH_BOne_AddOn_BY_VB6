VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY575"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY575.cls
'//  Module         : PH
'//  Desc           : 근태기찰현황
'//  작성자         : NGY
'//  DATE           : 2012.12.10
'//********************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form

''// 그리드 사용시
'Public oGrid1           As SAPbouiCOM.Grid
'Public oDS_PH_PY575     As SAPbouiCOM.DataTable
'
''// 매트릭스 사용시
'Public oMat1 As SAPbouiCOM.Matrix
'Private oDS_PH_PY575A As SAPbouiCOM.DBDataSource
'Private oDS_PH_PY575B As SAPbouiCOM.DBDataSource

Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long

Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY575.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY575_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY575"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
'    oForm.DataBrowser.BrowseBy = "Code"
        
    oForm.Freeze True
    Call PH_PY575_CreateItems
    Call PH_PY575_EnableMenus
    Call PH_PY575_SetDocument(oFromDocEntry01)
'    Call PH_PY575_FormResize
    
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

Private Function PH_PY575_CreateItems() As Boolean

    Dim sQry        As String
    Dim i           As Long
    
    Dim oCheck      As SAPbouiCOM.CheckBox
    Dim oEdit       As SAPbouiCOM.EditText
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns    As SAPbouiCOM.Columns
    Dim optBtn      As SAPbouiCOM.OptionBtn
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY575_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    
    '//----------------------------------------------------------------------------------------------
    '// 아이템 설정
    '//----------------------------------------------------------------------------------------------
    
    '//사업장
    Call oForm.DataSources.UserDataSources.Add("CLTCOD", dt_SHORT_TEXT, 10)
    Set oCombo = oForm.Items("CLTCOD").Specific
    oCombo.DataBind.SetBound True, "", "CLTCOD"
    
    oForm.Items("CLTCOD").DisplayDesc = True
    
    '// 접속자에 따른 권한별 사업장 콤보박스세팅
    Call CLTCOD_Select(oForm, "CLTCOD")
    
    '//기준일자
    Call oForm.DataSources.UserDataSources.Add("DocDateFr", dt_DATE, 10)
    oForm.Items("DocDateFr").Specific.DataBind.SetBound True, "", "DocDateFr"
    oForm.DataSources.UserDataSources.Item("DocDateFr").VALUE = Format(Now, "YYYYMM01")
    
    Call oForm.DataSources.UserDataSources.Add("DocDateTo", dt_DATE, 10)
    oForm.Items("DocDateTo").Specific.DataBind.SetBound True, "", "DocDateTo"
    oForm.DataSources.UserDataSources.Item("DocDateTo").VALUE = Format(Now, "YYYYMMDD")
    
    '//부서
    Call oForm.DataSources.UserDataSources.Add("TeamCode", dt_SHORT_TEXT, 10)
    Set oCombo = oForm.Items("TeamCode").Specific
    oCombo.DataBind.SetBound True, "", "TeamCode"
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '1' AND U_UseYN= 'Y' AND U_Char2 = '" & Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) & "'"
    Call SetReDataCombo(oForm, sQry, oCombo, "Y")
    oForm.Items("TeamCode").DisplayDesc = True
    
    '//담당
    Call oForm.DataSources.UserDataSources.Add("RspCode", dt_SHORT_TEXT, 10)
    Set oCombo = oForm.Items("RspCode").Specific
    oCombo.DataBind.SetBound True, "", "RspCode"
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '2' AND U_UseYN= 'Y' AND U_Char2 = '" & Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) & "'"
    Call SetReDataCombo(oForm, sQry, oCombo, "Y")
    oForm.Items("RspCode").DisplayDesc = True
    
    '반
    Call oForm.DataSources.UserDataSources.Add("ClsCode", dt_SHORT_TEXT, 20)
    Call oForm.Items("ClsCode").Specific.DataBind.SetBound(True, "", "ClsCode")
    
    '사번
    Call oForm.DataSources.UserDataSources.Add("MSTCOD", dt_SHORT_TEXT, 20)
    Call oForm.Items("MSTCOD").Specific.DataBind.SetBound(True, "", "MSTCOD")
    
    '성명
    Call oForm.DataSources.UserDataSources.Add("FullName", dt_SHORT_TEXT, 20)
    Call oForm.Items("FullName").Specific.DataBind.SetBound(True, "", "FullName")
    
    '//근무형태
    Call oForm.DataSources.UserDataSources.Add("ShiftDat", dt_SHORT_TEXT, 10)
    Set oCombo = oForm.Items("ShiftDat").Specific
    oCombo.DataBind.SetBound True, "", "ShiftDat"
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P154' AND U_UseYN= 'Y'"
    Call SetReDataCombo(oForm, sQry, oCombo, "Y")
    oCombo.Select 0, psk_Index
     
     '//근무조
    Call oForm.DataSources.UserDataSources.Add("GNMUJO", dt_SHORT_TEXT, 10)
    Set oCombo = oForm.Items("GNMUJO").Specific
    oCombo.DataBind.SetBound True, "", "GNMUJO"
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P155' AND U_UseYN= 'Y' AND U_Char1 = '" & oForm.Items.Item("ShiftDat").Specific.VALUE & "'"
    Call SetReDataCombo(oForm, sQry, oCombo, "Y")
    oCombo.Select 0, psk_Index
    
    '//구분
    Set oCombo = oForm.Items("Gubun").Specific
    oCombo.ValidValues.Add "%", "전체"
    oCombo.ValidValues.Add "Y", "근태이상"
    oCombo.ValidValues.Add "N", "정상"
    oCombo.Select 0, psk_Index
    
    '//근태확인완료구분
    Set oCombo = oForm.Items("Conf").Specific
    oCombo.ValidValues.Add "%", "전체"
    oCombo.ValidValues.Add "Y", "확인"
    oCombo.ValidValues.Add "N", "미확인"
    oCombo.Select 0, psk_Index
    
    'G5제외
    Call oForm.DataSources.UserDataSources.Add("G5_YN", dt_SHORT_TEXT, 1)
    Set oCheck = oForm.Items("G5_YN").Specific
    oCheck.DataBind.SetBound True, "", "G5_YN"
    oForm.Items("G5_YN").Specific.Checked = False
    
    '//커서를 첫번째 ITEM으로 지정
'    oForm.ActiveItem = "CLTCOD"
        
    
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
    
PH_PY575_CreateItems_Error:
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY575_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY575_EnableMenus()
    
    On Error GoTo PH_PY575_EnableMenus_Error

    Call oForm.EnableMenu("1283", True)     '//제거
    Call oForm.EnableMenu("1284", False)    '//취소
    Call oForm.EnableMenu("1293", True)     '//행삭제
    
    Exit Sub
    
PH_PY575_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY575_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY575_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY575_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY575_FormItemEnabled
        Call PH_PY575_AddMatrixRow
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY575_FormItemEnabled
        oForm.Items("Code").Specific.VALUE = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY575_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY575_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY575_FormItemEnabled()
    Dim oCombo      As SAPbouiCOM.ComboBox
    
    On Error GoTo PH_PY575_FormItemEnabled_Error
    
    
    Call oForm.Freeze(True)
    If (oForm.Mode = fm_ADD_MODE) Then

        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", False)    '//문서추가
        
    ElseIf (oForm.Mode = fm_FIND_MODE) Then

        Call oForm.EnableMenu("1281", False)    '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가
    ElseIf (oForm.Mode = fm_OK_MODE) Then

        
        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가

    End If
    Call oForm.Freeze(False)
    Exit Sub
    
PH_PY575_FormItemEnabled_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY575_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
                If pval.ItemUID = "Btn01" Then
                    Call PH_PY575_Print_Report01
                End If
            ElseIf pval.BeforeAction = False Then


            End If
     
'----------------------------------------------------------
        Case et_KEY_DOWN: '//2
        
            If pval.BeforeAction = True Then
            
                Call MDC_PS_Common.ActiveUserDefineValue(oForm, pval, BubbleEvent, "MSTCOD", "") '사번
                
            ElseIf pval.BeforeAction = False Then
            
            End If
'----------------------------------------------------------
        Case et_GOT_FOCUS: '//3
            Select Case pval.ItemUID
            Case "Mat01", "Grid01"
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
                            sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '1' AND U_UseYN= 'Y' AND U_Char2 = '" & Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) & "'"
                            Call SetReDataCombo(oForm, sQry, oCombo, "Y")
                            
                            '//담당
                            Set oCombo = oForm.Items("RspCode").Specific
                            If oCombo.ValidValues.Count > 0 Then              '//삭제
                               For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                                       oCombo.ValidValues.Remove i, psk_Index
                               Next i
                            End If
                            '//현재 사업장으로 다시 Qry
                            sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '2' AND U_UseYN= 'Y' AND U_Char2 = '" & Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) & "'"
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
                            
                            '반삭제
                            If oForm.Items("ClsCode").Specific.ValidValues.Count > 0 Then
                                For i = oForm.Items("ClsCode").Specific.ValidValues.Count - 1 To 0 Step -1
                                    Call oForm.Items("ClsCode").Specific.ValidValues.Remove(i, psk_Index)
                                Next i
                            End If
                            
                            '//현재 사업장으로 다시 Qry
                            sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '2' AND U_UseYN= 'Y' AND U_Char1 = '" & oForm.Items.Item("TeamCode").Specific.VALUE & "' AND U_Char2 = '" & Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) & "'"
                            Call SetReDataCombo(oForm, sQry, oCombo, "Y")
                        
                        '담당이 바뀌면 반 재설정
                        Case "RspCode"
                            
                            If oForm.Items("ClsCode").Specific.ValidValues.Count > 0 Then
                                For i = oForm.Items("ClsCode").Specific.ValidValues.Count - 1 To 0 Step -1
                                    Call oForm.Items("ClsCode").Specific.ValidValues.Remove(i, psk_Index)
                                Next i
                            End If
                            
                            '반콤보세팅
                            Call oForm.Items("ClsCode").Specific.ValidValues.Add("", "")
                            sQry = "            SELECT      U_Code AS [Code],"
                            sQry = sQry & "                 U_CodeNm As [Name]"
                            sQry = sQry & "  FROM       [@PS_HR200L]"
                            sQry = sQry & "  WHERE      Code = '9'"
                            sQry = sQry & "                 AND U_UseYN = 'Y'"
                            sQry = sQry & "                 AND U_Char1 = '" & Trim(oForm.Items("RspCode").Specific.VALUE) & "'"
                            sQry = sQry & "                 AND U_Char2 = '" & Trim(oForm.Items("TeamCode").Specific.VALUE) & "'"
                            sQry = sQry & "  ORDER BY  U_Seq"
                            Call MDC_SetMod.Set_ComboList(oForm.Items("ClsCode").Specific, sQry, "", False, False)
                            Call oForm.Items("ClsCode").Specific.Select(0, psk_Index)
                        
                        '//근무형태가 바뀌면 근무조 재설정
                        Case "ShiftDat"
                            Set oCombo = oForm.Items("GNMUJO").Specific
                            If oCombo.ValidValues.Count > 0 Then              '//삭제
                               For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                                       oCombo.ValidValues.Remove i, psk_Index
                               Next i
                            End If
                            sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P155' AND U_UseYN= 'Y' AND U_Char1 = '" & oForm.Items.Item("ShiftDat").Specific.VALUE & "'"
                            Call SetReDataCombo(oForm, sQry, oCombo, "Y")

                    End Select
                End If
            End If
            
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_CLICK: '//6
            If pval.BeforeAction = True Then
                Select Case pval.ItemUID
                Case "Mat01"
'                    If pval.Row > 0 Then
'                        Call oMat1.SelectRow(pval.Row, True, False)
'                    End If
                End Select
                
                Select Case pval.ItemUID
                Case "Mat01", "Grid01"
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
                            '//사원명 찿아서 화면 표시 하기
                            sQry = "SELECT U_FullName FROM [@PH_PY001A] WHERE Code =  '" & Trim(oForm.Items("MSTCOD").Specific.VALUE) & "'"
                            oRecordSet.DoQuery sQry
                            oForm.Items("FullName").Specific.String = Trim(oRecordSet.Fields("U_FullName").VALUE)
                    End Select
                End If
            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_MATRIX_LOAD: '//11
            If pval.BeforeAction = True Then
            ElseIf pval.BeforeAction = False Then
               'oMat1.LoadFromDataSource

                Call PH_PY575_FormItemEnabled
                Call PH_PY575_AddMatrixRow
                
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
'                Set oDS_PH_PY575A = Nothing
'                Set oDS_PH_PY575B = Nothing
                
                'Set oMat1 = Nothing
                'Set oGrid1 = Nothing
                
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
                Call PH_PY575_FormItemEnabled
                Call PH_PY575_AddMatrixRow
            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//문서찾기
                Call PH_PY575_FormItemEnabled
                Call PH_PY575_AddMatrixRow
                oForm.Items("Code").CLICK ct_Regular
            Case "1282":    '//문서추가
                Call PH_PY575_FormItemEnabled
                Call PH_PY575_AddMatrixRow
            Case "1288", "1289", "1290", "1291":
                Call PH_PY575_FormItemEnabled
            Case "1293" '// 행삭제
'                '// [MAT1 용]
'                 If oMat1.RowCount <> oMat1.VisualRowCount Then
'                    oMat1.FlushToDataSource
'
'                    While (i <= oDS_PH_PY575B.Size - 1)
'                        If oDS_PH_PY575B.GetValue("U_FILD01", i) = "" Then
'                            oDS_PH_PY575B.RemoveRecord (i)
'                            i = 0
'                        Else
'                            i = i + 1
'                        End If
'                    Wend
'
'                    For i = 0 To oDS_PH_PY575B.Size
'                        Call oDS_PH_PY575B.setValue("U_LineNum", i, i + 1)
'                    Next i
'
'                    oMat1.LoadFromDataSource
                'End If
                Call PH_PY575_AddMatrixRow
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

Sub PH_PY575_AddMatrixRow()
    Dim oRow    As Long
    
    On Error GoTo PH_PY575_AddMatrixRow_Error
    
    Call oForm.Freeze(True)
    
    '//[Mat1 용]
    'oMat1.FlushToDataSource
    'oRow = oMat1.VisualRowCount
'
'    If oMat1.VisualRowCount > 0 Then
'        If Trim(oDS_PH_PY575B.GetValue("U_FILD01", oRow - 1)) <> "" Then
'            If oDS_PH_PY575B.Size <= oMat1.VisualRowCount Then
'                oDS_PH_PY575B.InsertRecord (oRow)
'            End If
'            oDS_PH_PY575B.Offset = oRow
'            oDS_PH_PY575B.setValue "U_LineNum", oRow, oRow + 1
'            oDS_PH_PY575B.setValue "U_FILD01", oRow, ""
'            oDS_PH_PY575B.setValue "U_FILD02", oRow, ""
'            oDS_PH_PY575B.setValue "U_FILD03", oRow, 0
'            oMat1.LoadFromDataSource
'        Else
'            oDS_PH_PY575B.Offset = oRow - 1
'            oDS_PH_PY575B.setValue "U_LineNum", oRow - 1, oRow
'            oDS_PH_PY575B.setValue "U_FILD01", oRow - 1, ""
'            oDS_PH_PY575B.setValue "U_FILD02", oRow - 1, ""
'            oDS_PH_PY575B.setValue "U_FILD03", oRow - 1, 0
'            oMat1.LoadFromDataSource
'        End If
'    ElseIf oMat1.VisualRowCount = 0 Then
'        oDS_PH_PY575B.Offset = oRow
'        oDS_PH_PY575B.setValue "U_LineNum", oRow, oRow + 1
'        oDS_PH_PY575B.setValue "U_FILD01", oRow, ""
'        oDS_PH_PY575B.setValue "U_FILD02", oRow, ""
'        oDS_PH_PY575B.setValue "U_FILD03", oRow, 0
'        oMat1.LoadFromDataSource
'    End If
    
    Call oForm.Freeze(False)
    Exit Sub
PH_PY575_AddMatrixRow_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY575_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY575_FormClear()
On Error GoTo PH_PY575_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY575'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PH_PY575_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY575_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY575_DataValidCheck() As Boolean
On Error GoTo PH_PY575_DataValidCheck_Error
    PH_PY575_DataValidCheck = False
    Dim i       As Long
    Dim sQry    As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
'    '//----------------------------------------------------------------------------------
'    '//필수 체크
'    '//----------------------------------------------------------------------------------
'    If Trim(oDS_PH_PY575A.GetValue("Code", 0)) = "" Then
'        Sbo_Application.SetStatusBarMessage "사원번호는 필수입니다.", bmt_Short, True
'        oForm.Items("Code").CLICK ct_Regular
'        PH_PY575_DataValidCheck = False
'        Exit Function
'    End If
'
'    oMat1.FlushToDataSource
'    '// Matrix 마지막 행 삭제(DB 저장시)
'    If oDS_PH_PY575B.Size > 1 Then oDS_PH_PY575B.RemoveRecord (oDS_PH_PY575B.Size - 1)
'    oMat1.LoadFromDataSource
    
    PH_PY575_DataValidCheck = True
    
    Exit Function
    
    Set oRecordSet = Nothing
    
    
PH_PY575_DataValidCheck_Error:
    Set oRecordSet = Nothing
    PH_PY575_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY575_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Function PH_PY575_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY575_Validate_Error
    PH_PY575_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY575A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", bmt_Short, True
        PH_PY575_Validate = False
        GoTo PH_PY575_Validate_Exit
    End If
'
    If ValidateType = "수정" Then

    ElseIf ValidateType = "행삭제" Then

    ElseIf ValidateType = "취소" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY575_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY575_Validate_Error:
    PH_PY575_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY575_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Private Sub PH_PY575_Print_Report01()

    Dim DocNum          As String
    Dim ErrNum          As Integer
    Dim WinTitle        As String
    Dim ReportName      As String
    Dim sQry            As String
    Dim oRecordSet      As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY575_Print_Report01_Error
    
    Dim CLTCOD          As String
    Dim DocDateFr       As String
    Dim DocDateTo       As String
    Dim TeamCode        As String
    Dim RspCode         As String
    Dim ClsCode         As String
    Dim MSTCOD          As String
    Dim ShiftDat        As String
    Dim GNMUJO          As String
    Dim Gubun           As String
    Dim Conf            As String
    Dim G5_YN As String
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
    
     '/ ODBC 연결 체크
    If ConnectODBC = False Then
        GoTo PH_PY575_Print_Report01_Error
    End If
    
    
    '//인자 MOVE , Trim 시키기..
    CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE)
    DocDateFr = Trim(oForm.Items("DocDateFr").Specific.VALUE)
    DocDateTo = Trim(oForm.Items("DocDateTo").Specific.VALUE)
    TeamCode = Trim(oForm.Items("TeamCode").Specific.VALUE)
    RspCode = Trim(oForm.Items("RspCode").Specific.VALUE)
    ClsCode = Trim(oForm.Items("ClsCode").Specific.VALUE)
    MSTCOD = Trim(oForm.Items("MSTCOD").Specific.VALUE)
    ShiftDat = Trim(oForm.Items("ShiftDat").Specific.VALUE)
    GNMUJO = Trim(oForm.Items("GNMUJO").Specific.VALUE)
    Gubun = Trim(oForm.Items("Gubun").Specific.VALUE)
    Conf = Trim(oForm.Items("Conf").Specific.VALUE)
    
    If oForm.DataSources.UserDataSources.Item("G5_YN").VALUE = "Y" Then
        G5_YN = "Y"
    Else
        G5_YN = "N"
    End If
    
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/

    WinTitle = "[PH_PY575] 근태기찰현황"
    ReportName = "PH_PY575_01.rpt"
    
    '/ Formula 수식필드
    ReDim gRpt_Formula(4)
    ReDim gRpt_Formula_Value(4)
    
    gRpt_Formula(1) = "CLTCOD"
    sQry = "SELECT U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y' AND U_Code = '" & CLTCOD & "'"
    Call oRecordSet.DoQuery(sQry)
    gRpt_Formula_Value(1) = oRecordSet.Fields(0).VALUE
    
    gRpt_Formula(2) = "DocDateFr"
    gRpt_Formula_Value(2) = Format(DocDateFr, "####-##-##")
    
    gRpt_Formula(3) = "DocDateTo"
    gRpt_Formula_Value(3) = Format(DocDateTo, "####-##-##")
    
    gRpt_Formula(4) = "Gubun"
    gRpt_Formula_Value(4) = Gubun
    
    
    '/ SubReport
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)
    
    '/ Procedure 실행"
    sQry = "EXEC [PH_PY575_01] '" & CLTCOD & "', '" & DocDateFr & "',  '" & DocDateTo & "', '" & TeamCode & "', '" & RspCode & "', '" & ClsCode & "','" & MSTCOD & "', '" & ShiftDat & "', '" & GNMUJO & "', '" & Gubun & "', '" & Conf & "', '" & G5_YN & "'"
    
'    oRecordSet.DoQuery sQry
'    If oRecordSet.RecordCount = 0 Then
'        ErrNum = 1
'        GoTo PH_PY575_Print_Report01_Error
'    End If
    
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "1", "Y", "V", , 1) = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
    End If
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set oRecordSet = Nothing
    Exit Sub
    
PH_PY575_Print_Report01_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set oRecordSet = Nothing

    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
    Else
        Sbo_Application.SetStatusBarMessage "PH_PY575_Print_Report01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    End If
    
End Sub











