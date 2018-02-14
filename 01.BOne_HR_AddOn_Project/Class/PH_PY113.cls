VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY113"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY113.cls
'//  Module         : 급여관리 > 급(상)여 분개장 생성
'//  Desc           : 급(상)여 분개장 생성
'//********************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form

''// 매트릭스 사용시
Public oMat1 As SAPbouiCOM.Matrix

Private oDS_PH_PY113A As SAPbouiCOM.DBDataSource
Private oDS_PH_PY113B As SAPbouiCOM.DBDataSource

Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long

Private oCLTCOD         As String
Private oYM         As String
Private oJOBTYP         As String
Private oJOBGBN         As String
Private oPAYSEL         As String
Private oTOTDEB         As Double
Private oTOTCRE         As Double
Private oTOTPAY         As Double
Private oTOTGON         As Double

Private oDocDate        As String
Private oREMARK         As String
Private oDocNum         As String
Private oDIM3           As String

Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY113.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY113_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY113"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
    oForm.DataBrowser.BrowseBy = "DocEntry"
        
    oForm.Freeze True
    Call PH_PY113_CreateItems
    Call PH_PY113_EnableMenus
    Call PH_PY113_SetDocument(oFromDocEntry01)
'    Call PH_PY113_FormResize
    
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

Private Function PH_PY113_CreateItems() As Boolean

    Dim sQry        As String
    Dim i           As Long
    
    Dim oCheck      As SAPbouiCOM.CheckBox
    Dim oEdit       As SAPbouiCOM.EditText
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns    As SAPbouiCOM.Columns
    Dim optBtn      As SAPbouiCOM.OptionBtn
    Dim oLink       As SAPbouiCOM.LinkedButton
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY113_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    
    '//----------------------------------------------------------------------------------------------
    '// 데이터셋정의
    '//----------------------------------------------------------------------------------------------
'    '//테이블이 있을경우 데이터셋(Matrix)
    Set oDS_PH_PY113A = oForm.DataSources.DBDataSources("@PH_PY113A")   '//헤더
    Set oDS_PH_PY113B = oForm.DataSources.DBDataSources("@PH_PY113B")   '//라인

    Set oMat1 = oForm.Items("Mat1").Specific       '

    oMat1.SelectionMode = ms_NotSupported
    oMat1.AutoResizeColumns

    '//----------------------------------------------------------------------------------------------
    '// 아이템 설정
    '//----------------------------------------------------------------------------------------------
    '//사업장
    Set oCombo = oForm.Items("CLTCOD").Specific
    oForm.Items("CLTCOD").DisplayDesc = True
    
    '//지급종류
    Set oCombo = oForm.Items("JOBTYP").Specific
    oCombo.ValidValues.Add "1", "급여"
    oCombo.ValidValues.Add "2", "상여"
    oCombo.Select "0", psk_Index
    oForm.Items("JOBTYP").DisplayDesc = True
    
    '//지급구분
    Set oCombo = oForm.Items("JOBGBN").Specific
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P212' AND U_UseYN= 'Y'"
    Call SetReDataCombo(oForm, sQry, oCombo)
    oCombo.Select "0", psk_Index
    oForm.Items("JOBGBN").DisplayDesc = True
    
    
    '//지급대상자구분
    Set oCombo = oForm.Items("PAYSEL").Specific
    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P213' ORDER BY CAST(U_Code AS NUMERIC) "
    Call SetReDataCombo(oForm, sQry, oCombo)
    oCombo.ValidValues.Add "%", "전체"
    oCombo.Select "%", psk_ByValue
    oForm.Items("PAYSEL").DisplayDesc = True
    
    '// 전기일자
    Call oDS_PH_PY113A.setValue("U_DOCDATE", 0, Format(Date, "YYYYMMDD"))
    
    '/ Matrix
    Set oMat1 = oForm.Items("Mat1").Specific
    
    Set oColumn = oMat1.Columns("AcctCode")
    Set oLink = oColumn.ExtendedObject
    oLink.LinkedObject = lf_GLAccounts

    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oLink = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
    
PH_PY113_CreateItems_Error:
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oLink = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY113_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY113_EnableMenus()
    
    On Error GoTo PH_PY113_EnableMenus_Error

    Call oForm.EnableMenu("1283", False)     '//제거
    Call oForm.EnableMenu("1284", False)    '//취소
    Call oForm.EnableMenu("1293", True)     '//행삭제
    
    Exit Sub
    
PH_PY113_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY113_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY113_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY113_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY113_FormItemEnabled
        Call PH_PY113_AddMatrixRow
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY113_FormItemEnabled
        oForm.Items("DocEntry").Specific.Value = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY113_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY113_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY113_FormItemEnabled()
    Dim oCombo      As SAPbouiCOM.ComboBox
    
    On Error GoTo PH_PY113_FormItemEnabled_Error
    
    
    Call oForm.Freeze(True)
    If (oForm.Mode = fm_ADD_MODE) Then
        PH_PY113_FormClear
        oForm.ActiveItem = "CLTCOD"
        oForm.Items("DocEntry").Enabled = False
        
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        
        '// 귀속년월
        Call oDS_PH_PY113A.setValue("U_YM", 0, Format(Date, "YYYYMM"))
        '//지급종류
        Set oCombo = oForm.Items("JOBTYP").Specific
        oCombo.Select "1", psk_ByValue
        '//지급구분
        Set oCombo = oForm.Items("JOBGBN").Specific
        oCombo.Select "1", psk_ByValue
        '//지급대상자구분
        Set oCombo = oForm.Items("PAYSEL").Specific
        oCombo.Select "%", psk_ByValue
        
        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", False)    '//문서추가
        
    ElseIf (oForm.Mode = fm_FIND_MODE) Then
        oForm.Items("DocEntry").Enabled = True
        
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        
        Call oForm.EnableMenu("1281", False)    '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가
    ElseIf (oForm.Mode = fm_OK_MODE) Then
        oForm.Items("DocEntry").Enabled = False
        
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        
        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가

    End If
    Call oForm.Freeze(False)
    Exit Sub
    
PH_PY113_FormItemEnabled_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY113_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
                If pval.ItemUid = "1" Then
                    If PH_PY113_DataValidCheck = False Then
                        BubbleEvent = False
                    End If
                End If
                If pval.ItemUid = "Btn2" Then
                    If oForm.Mode = fm_ADD_MODE Then
                        If Execution_Process = False Then
                            BubbleEvent = False
                            Exit Sub
                        Else
                            BubbleEvent = False
                            Exit Sub
                        End If
                    Else
                        Sbo_Application.StatusBar.SetText "추가 모드에서만 조회가 가능합니다.", bmt_Short, smt_Error
                    End If
                End If
            ElseIf pval.BeforeAction = False Then
                If pval.ItemUid = "1" Then
                    If pval.ActionSuccess = True Then
                        Call PH_PY113_FormItemEnabled
                    End If
                End If
                
                If pval.ItemUid = "Btn1" Then    '/ 분개장 문서생성
                    If oDS_PH_PY113A.GetValue("DocEntry", 0) <> "" Then
                        If MDC_SetMod.Value_ChkYn("[@PH_PY113A]", "DocEntry", "'" & Trim(oDS_PH_PY113A.GetValue("DocEntry", 0)) & "'") = False Then
                            Call DI_oJournalEntries
                        Else
                            Sbo_Application.StatusBar.SetText "저장된 문서만 분개 생성이 가능합니다.", bmt_Short, smt_Error
                        End If
                    End If
                End If
            End If
     
'----------------------------------------------------------
        Case et_KEY_DOWN: '//2
            If pval.BeforeAction = True And pval.ColUID = "AcctCode" And pval.CharPressed = 9 Then
                If MDC_SetMod.Value_ChkYn("[OACT]", "FormatCode", "'" & oMat1.Columns(pval.ColUID).Cells(pval.Row).Specific.String & "'", "") = True Or _
                    oMat1.Columns(pval.ColUID).Cells(pval.Row).Specific.String = "" Then
                    Sbo_Application.ActivateMenuItem ("7425")
                    BubbleEvent = False
                End If
             End If
'----------------------------------------------------------
        Case et_GOT_FOCUS: '//3
            Select Case pval.ItemUid
            Case "Mat1", "Grid1"
                If pval.Row > 0 Then
                    oLastItemUID = pval.ItemUid
                    oLastColUID = pval.ColUID
                    oLastColRow = pval.Row
                End If
            Case Else
                oLastItemUID = pval.ItemUid
                oLastColUID = ""
                oLastColRow = 0
            End Select
'----------------------------------------------------------
        Case et_LOST_FOCUS: '//4

'----------------------------------------------------------
        Case et_COMBO_SELECT: '//5
'            Call oForm.Freeze(True)
'            If pval.BeforeAction = True Then
'
'            ElseIf pval.BeforeAction = False Then
'                If pval.ItemChanged = True Then
'
'                End If
'            End If
'            Call oForm.Freeze(False)
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
            If pval.BeforeAction = False And pval.ItemChanged = True Then
                If pval.ItemUid = "ENDDAT" Or pval.ItemUid = "MSTCOD" Then
                    If oForm.Items(pval.ItemUid).Specific.String = "" Then
                        oForm.Items(pval.ItemUid).Specific.Value = Format$(Now, "YYYYMMDD")
                    End If
                    oForm.Items(pval.ItemUid).Update
                End If
            ElseIf pval.BeforeAction = False And pval.ItemChanged = True Then
                If pval.ItemUid = "Mat1" Then
                    If (pval.ColUID = "AcctCode" Or pval.ColUID = "AcctName" Or _
                        pval.ColUID = "ShortNam" Or pval.ColUID = "Debit" Or pval.ColUID = "Credit") Then
                        
                        oMat1.FlushToDataSource
                        oDS_PH_PY113B.Offset = pval.Row - 1
                        Select Case pval.ColUID
                            Case "AcctCode"  '계정코드
                                oDS_PH_PY113B.setValue "U_Col04", pval.Row - 1, oMat1.Columns("AcctCode").Cells(pval.Row).Specific.Value
                                oDS_PH_PY113B.setValue "U_Col12", pval.Row - 1, MDC_SetMod.Get_ReData("AcctName", "AcctCode", "OACT", "'" & Trim$(oMat1.Columns("AcctCode").Cells(pval.Row).Specific.Value) & "'")
                            Case "Debit", "Credit"
                                oDS_PH_PY113B.setValue "U_Col05", pval.Row - 1, Format$(oMat1.Columns("Debit").Cells(pval.Row).Specific.Value, "#,###,###,##0")
                                oDS_PH_PY113B.setValue "U_Col06", pval.Row - 1, Format$(oMat1.Columns("Credit").Cells(pval.Row).Specific.Value, "#,###,###,##0")
                        End Select
                        oMat1.SetLineData pval.Row
                        Call TOTAL_AMT
                        oDS_PH_PY113B.Offset = pval.Row - 1
                        If pval.Row = oMat1.RowCount And Trim$(oDS_PH_PY113B.GetValue("U_AcctCode", pval.Row - 1)) <> "" Then
                            PH_PY113_AddMatrixRow
                            oMat1.Columns("AcctCode").Cells(pval.Row).CLICK ct_Regular
                        End If
                    End If
                End If
            End If
'----------------------------------------------------------
        Case et_MATRIX_LOAD: '//11
'            If pval.BeforeAction = True Then
'            ElseIf pval.BeforeAction = False Then
'                oMat1.LoadFromDataSource
'
'                Call PH_PY113_FormItemEnabled
'                Call PH_PY113_AddMatrixRow
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
                Set oDS_PH_PY113A = Nothing
                Set oDS_PH_PY113B = Nothing

                Set oMat1 = Nothing
                
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
                Call PH_PY113_FormItemEnabled
                Call PH_PY113_AddMatrixRow
            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//문서찾기
                Call PH_PY113_FormItemEnabled
                Call PH_PY113_AddMatrixRow
                oForm.Items("DocEntry").CLICK ct_Regular
            Case "1282":    '//문서추가
                Call PH_PY113_FormItemEnabled
                Call PH_PY113_AddMatrixRow
            Case "1288", "1289", "1290", "1291":
                Call PH_PY113_FormItemEnabled
            Case "1293" '// 행삭제
                '// [MAT1 용]
                 If oMat1.RowCount <> oMat1.VisualRowCount Then
                    oMat1.FlushToDataSource

                    While (i <= oDS_PH_PY113B.Size - 1)
                        If oDS_PH_PY113B.GetValue("U_AcctCode", i) = "" Then
                            oDS_PH_PY113B.RemoveRecord (i)
                            i = 0
                        Else
                            i = i + 1
                        End If
                    Wend

                    For i = 0 To oDS_PH_PY113B.Size
                        Call oDS_PH_PY113B.setValue("U_LineNum", i, i + 1)
                    Next i

                    oMat1.LoadFromDataSource
                End If
                Call PH_PY113_AddMatrixRow
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
    Select Case pval.ItemUid
    Case "Mat1"
        If pval.Row > 0 Then
            oLastItemUID = pval.ItemUid
            oLastColUID = pval.ColUID
            oLastColRow = pval.Row
        End If
    Case Else
        oLastItemUID = pval.ItemUid
        oLastColUID = ""
        oLastColRow = 0
    End Select
    Exit Sub
Raise_RightClickEvent_Error:

    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY113_AddMatrixRow()
    Dim oRow    As Long
    
    On Error GoTo PH_PY113_AddMatrixRow_Error
    
    Call oForm.Freeze(True)
    
    '//[Mat1 용]
    oMat1.FlushToDataSource
    oRow = oMat1.VisualRowCount

    If oMat1.VisualRowCount > 0 Then
        If Trim(oDS_PH_PY113B.GetValue("U_AcctCode", oRow - 1)) <> "" Then
            If oDS_PH_PY113B.Size <= oMat1.VisualRowCount Then
                oDS_PH_PY113B.InsertRecord (oRow)
            End If
            oDS_PH_PY113B.Offset = oRow
            oDS_PH_PY113B.setValue "U_LineNum", oRow, oRow + 1
            oDS_PH_PY113B.setValue "U_AcctCode", oRow, ""
            oDS_PH_PY113B.setValue "U_AcctName", oRow, ""
            oDS_PH_PY113B.setValue "U_ShortNam", oRow, ""
            oDS_PH_PY113B.setValue "U_Debit", oRow, 0
            oDS_PH_PY113B.setValue "U_Credit", oRow, 0
            oDS_PH_PY113B.setValue "U_Project", oRow, ""
            oDS_PH_PY113B.setValue "U_CostCent", oRow, ""
            oDS_PH_PY113B.setValue "U_Comments", oRow, ""
            oMat1.LoadFromDataSource
        Else
            oDS_PH_PY113B.Offset = oRow - 1
            oDS_PH_PY113B.setValue "U_LineNum", oRow - 1, oRow
            oDS_PH_PY113B.setValue "U_AcctCode", oRow - 1, ""
            oDS_PH_PY113B.setValue "U_AcctName", oRow - 1, ""
            oDS_PH_PY113B.setValue "U_ShortNam", oRow - 1, ""
            oDS_PH_PY113B.setValue "U_Debit", oRow - 1, 0
            oDS_PH_PY113B.setValue "U_Credit", oRow - 1, 0
            oDS_PH_PY113B.setValue "U_Project", oRow - 1, ""
            oDS_PH_PY113B.setValue "U_CostCent", oRow - 1, ""
            oDS_PH_PY113B.setValue "U_Comments", oRow - 1, ""
            oMat1.LoadFromDataSource
        End If
    ElseIf oMat1.VisualRowCount = 0 Then
        oDS_PH_PY113B.Offset = oRow
        oDS_PH_PY113B.setValue "U_LineNum", oRow, oRow + 1
        oDS_PH_PY113B.setValue "U_AcctCode", oRow, ""
        oDS_PH_PY113B.setValue "U_AcctName", oRow, ""
        oDS_PH_PY113B.setValue "U_ShortNam", oRow, ""
        oDS_PH_PY113B.setValue "U_Debit", oRow, 0
        oDS_PH_PY113B.setValue "U_Credit", oRow, 0
        oDS_PH_PY113B.setValue "U_Project", oRow, ""
        oDS_PH_PY113B.setValue "U_CostCent", oRow, ""
        oDS_PH_PY113B.setValue "U_Comments", oRow, ""
        oMat1.LoadFromDataSource
    End If
    
    Call oForm.Freeze(False)
    Exit Sub
PH_PY113_AddMatrixRow_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY113_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY113_FormClear()
On Error GoTo PH_PY113_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY113'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.Value = 1
    Else
        oForm.Items("DocEntry").Specific.Value = DocEntry
    End If
    Exit Sub
PH_PY113_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY113_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY113_DataValidCheck() As Boolean
On Error GoTo PH_PY113_DataValidCheck_Error
    
    Dim i       As Long
    Dim sQry    As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    PH_PY113_DataValidCheck = False
    
    If oForm.Items("DOCNUM").Specific.Value <> "" Then
        If Sbo_Application.MessageBox("이미 분개된 자료입니다. 다시 분개하시겠습니까?", 1, "예", "아니오") = 2 Then
            Exit Function
        End If
    End If
    
    oDocDate = oForm.Items("DOCDATE").Specific.Value
    oREMARK = oForm.Items("REMARK").Specific.Value
    
    '//----------------------------------------------------------------------------------
    '//필수 체크
    '//----------------------------------------------------------------------------------
    If Trim(oDS_PH_PY113A.GetValue("U_DOCDATE", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "전기일은 필수입니다.", bmt_Short, True
        oForm.Items("DOCDATE").CLICK ct_Regular
        Exit Function
    End If

    If Trim(oDS_PH_PY113A.GetValue("U_TOTDEB", 0)) = 0 Then
        Sbo_Application.SetStatusBarMessage "차변합계가 0입니다.", bmt_Short, True
        oForm.Items("TOTDEB").CLICK ct_Regular
        Exit Function
    End If
    
    If Trim(oDS_PH_PY113A.GetValue("U_TOTCRE", 0)) = 0 Then
        Sbo_Application.SetStatusBarMessage "대변합계가 0입니다. ", bmt_Short, True
        oForm.Items("TOTCRE").CLICK ct_Regular
        Exit Function
    End If

    If Trim(oDS_PH_PY113A.GetValue("U_TOTDEB", 0)) <> Trim(oDS_PH_PY113A.GetValue("U_TOTCRE", 0)) Then
        Sbo_Application.SetStatusBarMessage "차변과 대변금액이 일치하지 않습니다.", bmt_Short, True
        oForm.Items("Code").CLICK ct_Regular
        Exit Function
    End If
    
    If Trim(oDS_PH_PY113A.GetValue("U_TOTDEB", 0)) <> Trim(oDS_PH_PY113A.GetValue("U_TOTPAY", 0)) Then
        Sbo_Application.SetStatusBarMessage "차변과 총지급액이 일치하지 않습니다.", bmt_Short, True
        oForm.Items("Code").CLICK ct_Regular
        Exit Function
    End If
    
    If oMat1.VisualRowCount > 1 Then
        For i = 1 To oMat1.VisualRowCount - 1
            If oMat1.Columns("U_AcctCode").Cells(i).Specific.Value = "" Then
                Sbo_Application.SetStatusBarMessage "계정과목코드 필수입니다", bmt_Short, True
                oMat1.Columns("U_AcctCode").Cells(i).CLICK ct_Regular
                Exit Function
            End If
            
            If Val(oMat1.Columns("U_Debit").Cells(i).Specific.Value) = 0 Then
                Sbo_Application.SetStatusBarMessage "차변금액이 0 입니다", bmt_Short, True
                oMat1.Columns("U_Debit").Cells(i).CLICK ct_Regular
                Exit Function
            End If
            
            If Val(oMat1.Columns("U_Credit").Cells(i).Specific.Value) = 0 Then
                Sbo_Application.SetStatusBarMessage "대변금액이 0 입니다", bmt_Short, True
                oMat1.Columns("U_Credit").Cells(i).CLICK ct_Regular
                Exit Function
            End If
        Next i
    Else
        Sbo_Application.SetStatusBarMessage "분개 자료가 없습니다.", bmt_Short, True
        Exit Function
    End If
    
    
    oMat1.FlushToDataSource
    '// Matrix 마지막 행 삭제(DB 저장시)
    If oDS_PH_PY113B.Size > 1 Then oDS_PH_PY113B.RemoveRecord (oDS_PH_PY113B.Size - 1)
    oMat1.LoadFromDataSource
    
    PH_PY113_DataValidCheck = True
    
    Set oRecordSet = Nothing
    Exit Function
    
PH_PY113_DataValidCheck_Error:
    Set oRecordSet = Nothing
    PH_PY113_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY113_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Function PH_PY113_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY113_Validate_Error
    PH_PY113_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY113A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.Value & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", bmt_Short, True
        PH_PY113_Validate = False
        GoTo PH_PY113_Validate_Exit
    End If
'
    If ValidateType = "수정" Then

    ElseIf ValidateType = "행삭제" Then

    ElseIf ValidateType = "취소" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY113_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY113_Validate_Error:
    PH_PY113_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY113_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function



Private Sub PH_PY113_Print_Report01()

    Dim DocNum          As String
    Dim ErrNum          As Integer
    Dim WinTitle        As String
    Dim ReportName      As String
    Dim sQry            As String
    
    Dim BPLID           As String
    Dim ItmBsort        As String
    Dim DocDate         As String
    
    Dim oRecordSet      As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY113_Print_Report01_Error
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '/ ODBC 연결 체크
    If ConnectODBC = False Then
        GoTo PH_PY113_Print_Report01_Error
    End If
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/

    WinTitle = "[S142] 발주서"
    ReportName = "S142_1.rpt"
    sQry = "EXEC PH_PY113_1 '" & oForm.Items("8").Specific.Value & "'"
    
    '/ Formula 수식필드
    ReDim gRpt_Formula(1)
    ReDim gRpt_Formula_Value(1)
    
    '/ SubReport
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)
    
    '/ Procedure 실행"
    sQry = "EXEC [PS_PP820_01] '" & BPLID & "', '" & ItmBsort & "', '" & DocDate & "'"
    
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
        If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "1", "Y", "V") = False Then
            Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
        End If
    Else
        Sbo_Application.SetStatusBarMessage "조회된 데이터가 없습니다.", bmt_Short, True
    End If
    
    Exit Sub
    
PH_PY113_Print_Report01_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY113_Print_Report01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub



Private Sub DI_oJournalEntries()
On Error GoTo Error_Message
    Dim f_oJournalEntries   As SAPbobsCOM.JournalVouchers
    Dim oRecordSet          As SAPbobsCOM.Recordset
    Dim sQry        As String
    Dim RetVal      As Long
    Dim nErr        As Long
    Dim ErrMsg      As String
    Dim i&, k&
    Dim AcctCode    As String
    Dim shortCode   As String
    Dim Credit      As Double
    Dim Debit       As Double
    Dim LineMemo    As String
    Dim Project     As String
    Dim Dimenz1     As String
    Dim Dimenz2     As String
    Dim Dimenz3     As String
    
    If Trim$(oDocDate) = "" Then
        oDocDate = Format$(Now, "yyyy-mm-dd")
    Else
        oDocDate = Format$(oDocDate, "0000-00-00")
    End If
    '/// 분개장문서
'    Set f_oJournalEntries = oCompany.GetBusinessObject(oJournalEntries)
    Set f_oJournalEntries = oCompany.GetBusinessObject(oJournalVouchers)
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '/ 재무관리>분개 =계정정보
    oCompany.StartTransaction
    With f_oJournalEntries
    '// 전표전체적용
'        .Series = "14"                             '/ 시리즈:주요 분개개체번호
'        .Series = MDC_Setmod.Get_Series_No("30")    '시리즈:주요 분개개체번호
        .JournalEntries.DueDate = oDocDate          '"04/26/2007"   '/ 만기일
        .JournalEntries.TaxDate = oDocDate          '"04/26/2007"   '/ 과세일
        .JournalEntries.ReferenceDate = oDocDate    '"04/26/2007"   '/ 전기일
        .JournalEntries.Memo = Trim$(oREMARK)
'        .ProjectCode = "103"                       '/ 프로젝트코드
'        .AutoVAT = tYES                            '/ 부가세코드사용
        i = 1
        oMat1.FlushToDataSource
        For k = 0 To oMat1.VisualRowCount - 1
            oDS_PH_PY113B.Offset = k
            If Trim$(oDS_PH_PY113B.GetValue("U_AcctCode", k)) <> "" Then
                If i <> 1 Then
                    .JournalEntries.Lines.Add
                    .JournalEntries.Lines.SetCurrentLine (k)
                End If
                AcctCode = oDS_PH_PY113B.GetValue("U_AcctCode", k)
                shortCode = Trim$(oDS_PH_PY113B.GetValue("U_ShortNam", k))
                Debit = Val(Replace(oDS_PH_PY113B.GetValue("U_Debit", k), ",", ""))
                Credit = Val(Replace(oDS_PH_PY113B.GetValue("U_Credit", k), ",", ""))
                LineMemo = Trim$(oDS_PH_PY113B.GetValue("U_Comments", k))
                Project = Trim$(oDS_PH_PY113B.GetValue("U_Prject", k))
'                Dimenz1 = Trim$(oDS_PH_PY113B.GetValue("U_Col07", k))
'                Dimenz2 = Trim$(oDS_PH_PY113B.GetValue("U_Col08", k))
'                If oDIM3 = "Y" Then Dimenz3 = Trim$(oDS_PH_PY113B.GetValue("U_Col10", k))
                
              '  .Lines.SetCurrentLine (0)
                .JournalEntries.Lines.AccountCode = MDC_SetMod.Get_ReData("AcctCode", "FormatCode", "[OACT]", "'" & Trim$(AcctCode) & "'")
                If MDC_SetMod.Value_ChkYn("[OACT]", "FormatCode", "'" & Trim$(shortCode) & "'") = False Then
                    .JournalEntries.Lines.ShortName = MDC_SetMod.Get_ReData("AcctCode", "FormatCode", "[OACT]", "'" & Trim$(shortCode) & "'")
                Else
                    .JournalEntries.Lines.ShortName = Trim$(shortCode)
                End If
                .JournalEntries.Lines.Credit = Credit
                .JournalEntries.Lines.Debit = Debit
                .JournalEntries.Lines.ProjectCode = Trim$(Project)
'                .JournalEntries.Lines.CostingCode = Trim$(Dimenz1)  '차원1
'                .JournalEntries.Lines.CostingCode2 = Trim$(Dimenz2) '차원2
                .JournalEntries.Lines.LineMemo = Trim$(LineMemo)
'                If oDIM3 = "Y" Then
'                    .JournalEntries.Lines.CostingCode3 = Trim$(Dimenz3)  '차원3
'                End If
                i = i + 1
            End If
        Next k
        RetVal = .Add

    End With
    'Check Error
    If (0 <> RetVal) Then
        Call oCompany.GetLastError(nErr, ErrMsg)
        GoTo Error_Message  '//저장시 에러 발생
    End If
    oCompany.GetNewObjectCode oDocNum
     
    sQry = "EXEC PH_PY113_INSERT '" & Trim$(oDocNum) & "', 'PH_PY113', '" & Trim$(oDocDate) & "', " & _
                                    oTOTDEB & ", " & oTOTCRE & ", '" & _
                                    Trim$(oYM) & "', '" & Trim$(oJOBGBN) & "', '" & _
                                    Trim$(oJOBTYP) & "', '" & Trim$(oPAYSEL) & "', '" & _
                                    Trim$(oCLTCOD) & "'"
                                    
    oRecordSet.DoQuery sQry
     
    Call oCompany.EndTransaction(wf_Commit)
    oForm.Items("DOCNUM").Specific.Value = Left$(oDocNum, Len(Trim$(oDocNum)) - 1) & "-" & Right$(oDocNum, 1)
    oForm.Items("DOCNUM").Update
    
    '/ MsgBox ("완료!")
    Sbo_Application.StatusBar.SetText "분개장 문서가 생성되었습니다.", bmt_Short, smt_Success
    Set f_oJournalEntries = Nothing
    Exit Sub
'// Message /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
Error_Message:
    If oCompany.InTransaction Then
       oCompany.EndTransaction (wf_RollBack)
    End If
    Set f_oJournalEntries = Nothing
    Sbo_Application.StatusBar.SetText "oJournalEntries Error : " & Space$(10) & nErr & "," & ErrMsg, bmt_Short, smt_Error
End Sub

Private Function Execution_Process() As Boolean
On Error GoTo Error_Message
    Dim oRecordSet  As SAPbobsCOM.Recordset
    Dim sQry        As String
    Dim ErrNum      As Integer
    Dim i&
'/ Check
    ErrNum = 0
    oYM = oForm.Items("YM").Specific.Value
    If Trim$(oYM) = "" Then
        ErrNum = 1
        GoTo Error_Message
    End If
    oCLTCOD = oDS_PH_PY113A.GetValue("U_CLTCOD", 0)
    oJOBTYP = oDS_PH_PY113A.GetValue("U_JOBTYP", 0)
    oJOBGBN = oDS_PH_PY113A.GetValue("U_JOBGBN", 0)
    oPAYSEL = oDS_PH_PY113A.GetValue("U_PAYSEL", 0)
    
    '/ 초기화
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    oDS_PH_PY113B.Clear
    oMat1.LoadFromDataSource
    i = 0
    sQry = " EXEC PH_PY113 '" & Trim$(oCLTCOD) & "','" & Trim$(oYM) & "', '" & Trim$(oJOBTYP) & "', '"
    sQry = sQry & Trim$(oJOBGBN) & "', '" & Trim$(oPAYSEL) & "'"

    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
        oForm.Items("DOCNUM").Specific.Value = ""
        ErrNum = 2
        GoTo Error_Message
        
    End If
    oTOTDEB = 0: oTOTCRE = 0
    Do Until oRecordSet.EOF
        oDS_PH_PY113B.InsertRecord (i)
        oDS_PH_PY113B.Offset = i
        oDS_PH_PY113B.setValue "U_LineNum", i, i + 1
        oDS_PH_PY113B.setValue "U_AcctCode", i, oRecordSet.Fields("AcctCode").Value
        oDS_PH_PY113B.setValue "U_AcctName", i, oRecordSet.Fields("AcctName").Value
        oDS_PH_PY113B.setValue "U_ShortNam", i, oRecordSet.Fields("ShortNam").Value
        oDS_PH_PY113B.setValue "U_Debit", i, Format$(oRecordSet.Fields("Debit").Value, "#,###,###,##0")  '/ 총급여액
        oDS_PH_PY113B.setValue "U_Credit", i, Format$(oRecordSet.Fields("Credit").Value, "#,###,###,##0")
        oDS_PH_PY113B.setValue "U_Project", i, oRecordSet.Fields("U_Project").Value
        oDS_PH_PY113B.setValue "U_CostCent", i, oRecordSet.Fields("U_PNLCOD").Value
        oDS_PH_PY113B.setValue "U_Comments", i, oRecordSet.Fields("U_Remark").Value
    
        oTOTDEB = oTOTDEB + oRecordSet.Fields("Debit").Value
        oTOTCRE = oTOTCRE + oRecordSet.Fields("Credit").Value
        i = i + 1
        oRecordSet.MoveNext
    Loop
    oMat1.LoadFromDataSource
    
    Call PH_PY113_AddMatrixRow
    
    '/ 분개No 조회
    sQry = "EXEC PH_PY113_QUERY 'PH_PY113', '" & Trim$(oCLTCOD) & "','" & Trim$(oYM) & "', '"
    sQry = sQry & Trim$(oJOBGBN) & "', '" & Trim$(oJOBTYP) & "', '" & Trim$(oPAYSEL) & "'"
                              
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        oDocNum = oRecordSet.Fields("DOCNUM").Value
    Else
        oDocNum = ""
    End If
    oForm.Items("DOCNUM").Specific.Value = oDocNum
    
    
    oForm.Items("TOTPAY").Specific.Value = oTOTDEB
    oForm.Items("TOTGON").Specific.Value = oTOTCRE
    oForm.Items("TOTDEB").Specific.Value = oTOTDEB
    oForm.Items("TOTCRE").Specific.Value = oTOTCRE

 '/ End
 
    Sbo_Application.StatusBar.SetText "작업을 완료하였습니다.", bmt_Short, smt_Success
    Execution_Process = True
    Set oRecordSet = Nothing
    Exit Function
'/////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set oRecordSet = Nothing
    
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "조회월을 입력하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 2 Then
        Sbo_Application.StatusBar.SetText "조건과 일치하는 자료가 없습니다.", bmt_Short, smt_Error
    ElseIf ErrNum = 3 Then
        Sbo_Application.StatusBar.SetText "자사코드가 없습니다. 선택하여 주십시오.", bmt_Short, smt_Error
    Else
        Sbo_Application.StatusBar.SetText "Execution_Process 실행 중 오류가 발생했습니다." & Space$(10) & Err.Description, bmt_Short, smt_Error
    End If
    Execution_Process = False
End Function


Private Sub TOTAL_AMT()
Dim oRow   As Integer
    oTOTDEB = 0
    oTOTCRE = 0
    For oRow = 1 To oMat1.VisualRowCount
        oDS_PH_PY113B.Offset = oRow - 1
        oTOTDEB = oTOTDEB + Val(Replace(oDS_PH_PY113B.GetValue("U_Debit", oRow - 1), ",", ""))
        oTOTCRE = oTOTCRE + Val(Replace(oDS_PH_PY113B.GetValue("U_Credit", oRow - 1), ",", ""))
    Next oRow
 '/
    oForm.Items("TOTDEB").Specific.Value = oTOTDEB
    oForm.Items("TOTCRE").Specific.Value = oTOTCRE
End Sub


