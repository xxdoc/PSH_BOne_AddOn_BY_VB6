VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY122"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY122.cls
'//  Module         : 인사관리 > 급여관리
'//  Desc           : 급상여출력 개인부서설정등록
'//********************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form

Public oMat1 As SAPbouiCOM.Matrix

Private oDS_PH_PY122A As SAPbouiCOM.DBDataSource
Private oDS_PH_PY122B As SAPbouiCOM.DBDataSource

Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long

Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY122.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY122_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY122"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
    oForm.DataBrowser.BrowseBy = "Code"
    
    oForm.Freeze True
    Call PH_PY122_CreateItems
    Call PH_PY122_EnableMenus
    Call PH_PY122_SetDocument(oFromDocEntry01)
'    Call PH_PY122_FormResize
    
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

Private Function PH_PY122_CreateItems() As Boolean

    Dim sQry        As String
    Dim i           As Long
    Dim CLTCOD      As String
    

    Dim oEdit       As SAPbouiCOM.EditText
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns    As SAPbouiCOM.Columns

    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY122_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Set oDS_PH_PY122A = oForm.DataSources.DBDataSources("@PH_PY122A")
    Set oDS_PH_PY122B = oForm.DataSources.DBDataSources("@PH_PY122B")

    Set oMat1 = oForm.Items("Mat1").Specific       '//@PH_PY122B

    oMat1.SelectionMode = ms_NotSupported
    oMat1.AutoResizeColumns

    '//사업장
    Set oCombo = oForm.Items("CLTCOD").Specific
'    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
'    Call SetReDataCombo(oForm, sQry, oCombo)
    oForm.Items("CLTCOD").DisplayDesc = True
    
    CLTCOD = MDC_SetMod.Get_ReData("Branch", "USER_CODE", "OUSR", "'" & oCompany.UserName & "'")
    oCombo.Select CLTCOD, psk_ByValue
    
    oForm.Items("CLTCOD").DisplayDesc = True
    
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing

    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
    
PH_PY122_CreateItems_Error:
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing

    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY122_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY122_EnableMenus()
    
    On Error GoTo PH_PY122_EnableMenus_Error

    Call oForm.EnableMenu("1283", True)     '//제거
    Call oForm.EnableMenu("1284", False)    '//취소
    Call oForm.EnableMenu("1293", True)     '//행삭제
    
    Exit Sub
    
PH_PY122_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY122_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY122_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY122_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY122_FormItemEnabled
        Call PH_PY122_AddMatrixRow
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY122_FormItemEnabled
        oForm.Items("Code").Specific.VALUE = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY122_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY122_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY122_FormItemEnabled()
    Dim oCombo      As SAPbouiCOM.ComboBox
    
    On Error GoTo PH_PY122_FormItemEnabled_Error
    
    
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
    
PH_PY122_FormItemEnabled_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY122_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Public Sub Raise_FormItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
    Dim sQry        As String
    Dim i           As Long
    Dim tSex        As String
    Dim tBrith      As String
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo Raise_FormItemEvent_Error
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Select Case pval.EventType
        Case et_ITEM_PRESSED: '//1
            If pval.BeforeAction = True Then
                If pval.ItemUID = "1" Then
                
                    If pval.FormMode = fm_ADD_MODE Then
                    
                        If PH_PY122_DataValidCheck = False Then
                            BubbleEvent = False
                        End If
                
                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
                    ElseIf oForm.Mode = fm_OK_MODE Then
                    End If
                    
                End If
            ElseIf pval.BeforeAction = False Then
                If pval.ItemUID = "1" Then
                    If pval.ActionSuccess = True Then
                        Call PH_PY122_FormItemEnabled
                        Call PH_PY122_AddMatrixRow
                    End If
                End If
            End If
'----------------------------------------------------------
        Case et_KEY_DOWN: '//2
            If pval.BeforeAction = True Then
                If pval.CharPressed = 9 Then
                    If pval.ItemUID = "Mat1" Then
                        If pval.ColUID = "MSTCOD" Then
                            If oMat1.Columns("MSTCOD").Cells(pval.Row).Specific.VALUE = "" Then
                                Sbo_Application.ActivateMenuItem ("7425")
                                BubbleEvent = False
                            End If
                        ElseIf pval.ColUID = "ATeamCod" Then
                            If oMat1.Columns("ATeamCod").Cells(pval.Row).Specific.VALUE = "" Then
                                Sbo_Application.ActivateMenuItem ("7425")
                                BubbleEvent = False
                            End If
                        ElseIf pval.ColUID = "ARspCode" Then
                            If oMat1.Columns("ARspCode").Cells(pval.Row).Specific.VALUE = "" Then
                                Sbo_Application.ActivateMenuItem ("7425")
                                BubbleEvent = False
                            End If
                        End If
                    End If
                End If
            ElseIf pval.BeforeAction = False Then
                
            End If
'----------------------------------------------------------
        Case et_GOT_FOCUS: '//3
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
                Case "Mat1"
                    If pval.Row > 0 Then
                        Call oMat1.SelectRow(pval.Row, True, False)
                    End If
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
                
                If pval.ItemChanged = True Then
                    If pval.ColUID = "MSTCOD" Then
                        
                        oMat1.FlushToDataSource
                        
                        sQry = "select a.Code, FullName = a.U_FullName, TeamCode = a.U_TeamCode, TeamName = (Select U_CodeNm From [@PS_HR200L] Where Code = '1' and U_Code = a.U_TeamCode) "
                        sQry = sQry + " From [@PH_PY001A] a "
                        sQry = sQry + " Where a.U_CLTCOD = '" + oForm.Items("CLTCOD").Specific.VALUE + "'"
                        sQry = sQry + " and a.Code = '" + RTrim(oDS_PH_PY122B.GetValue("U_MSTCOD", pval.Row - 1)) + "'"
                        
                        oRecordSet.DoQuery sQry
                        
                        If oRecordSet.RecordCount = 0 Then
                            Sbo_Application.SetStatusBarMessage "선택한 사원이 없습니다.", bmt_Short, True
                        Else
                            Do Until oRecordSet.EOF
                                oDS_PH_PY122B.setValue "U_MSTNAM", pval.Row - 1, Trim(oRecordSet.Fields("Fullname").VALUE)
                                oDS_PH_PY122B.setValue "U_TeamCode", pval.Row - 1, Trim(oRecordSet.Fields("TeamCode").VALUE)
                                oDS_PH_PY122B.setValue "U_TeamName", pval.Row - 1, Trim(oRecordSet.Fields("TeamName").VALUE)
                                
                                oRecordSet.MoveNext
                            Loop
                        End If
                        oMat1.LoadFromDataSource
                        Call PH_PY122_AddMatrixRow
                        Call oMat1.Columns(pval.ColUID).Cells(pval.Row).CLICK(ct_Regular)
                    ElseIf pval.ColUID = "ATeamCod" Then
                        oMat1.FlushToDataSource
                        
                        sQry = "select a.U_CodeNm "
                        sQry = sQry + " From [@PS_HR200L] a "
                        sQry = sQry + " Where a.Code = '1' "
                        sQry = sQry + " and a.U_Code = '" + RTrim(oDS_PH_PY122B.GetValue("U_AteamCod", pval.Row - 1)) + "'"
                        
                        oRecordSet.DoQuery sQry
                        
                        If oRecordSet.RecordCount = 0 Then
                            Sbo_Application.SetStatusBarMessage "선택한 부서가 없습니다.", bmt_Short, True
                        Else
                            Do Until oRecordSet.EOF
                                oDS_PH_PY122B.setValue "U_ATeamNam", pval.Row - 1, Trim(oRecordSet.Fields("U_CodeNm").VALUE)
                                oRecordSet.MoveNext
                            Loop
                        End If
                        oMat1.LoadFromDataSource
                    ElseIf pval.ColUID = "ARspCode" Then
                        oMat1.FlushToDataSource
                        
                        sQry = "select a.U_CodeNm "
                        sQry = sQry + " From [@PS_HR200L] a "
                        sQry = sQry + " Where a.Code = '2' "
                        sQry = sQry + " and a.U_Code = '" + RTrim(oDS_PH_PY122B.GetValue("U_ARspCode", pval.Row - 1)) + "'"
                        
                        oRecordSet.DoQuery sQry
                        
                        If oRecordSet.RecordCount = 0 Then
                            Sbo_Application.SetStatusBarMessage "선택한 담당이 없습니다.", bmt_Short, True
                        Else
                            Do Until oRecordSet.EOF
                                oDS_PH_PY122B.setValue "U_ARspName", pval.Row - 1, Trim(oRecordSet.Fields("U_CodeNm").VALUE)
                                oRecordSet.MoveNext
                            Loop
                        End If
                        oMat1.LoadFromDataSource
                    End If
                End If
            ElseIf pval.BeforeAction = False Then
            End If
            
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_MATRIX_LOAD: '//11
            If pval.BeforeAction = True Then
            ElseIf pval.BeforeAction = False Then
                oMat1.LoadFromDataSource

                Call PH_PY122_FormItemEnabled
                Call PH_PY122_AddMatrixRow
                
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
                Set oDS_PH_PY122A = Nothing
                Set oDS_PH_PY122B = Nothing
                
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
                Call AuthorityCheck(oForm, "CLTCOD", "@PH_PY122A", "Code")      '//접속자 권한에 따른 사업장 보기
        End Select
    ElseIf (pval.BeforeAction = False) Then
        Select Case pval.MenuUID
            Case "1283":
                oForm.Mode = fm_ADD_MODE
                Call PH_PY122_FormItemEnabled
                Call PH_PY122_AddMatrixRow
            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//문서찾기
                Call PH_PY122_FormItemEnabled
                Call PH_PY122_AddMatrixRow
                
            Case "1282":    '//문서추가
                Call PH_PY122_FormItemEnabled
                Call PH_PY122_AddMatrixRow
            Case "1288", "1289", "1290", "1291":
                Call PH_PY122_FormItemEnabled
            Case "1293" '// 행삭제
                 If oMat1.RowCount <> oMat1.VisualRowCount Then
                    oMat1.FlushToDataSource
    
                    While (i <= oDS_PH_PY122B.Size - 1)
                        If oDS_PH_PY122B.GetValue("U_MSTCOD", i) = "" Then
                            oDS_PH_PY122B.RemoveRecord (i)
                            i = 0
                        Else
                            i = i + 1
                        End If
                    Wend
                    
                    For i = 0 To oDS_PH_PY122B.Size
                        Call oDS_PH_PY122B.setValue("U_LineNum", i, i + 1)
                    Next i
                    
                    oMat1.LoadFromDataSource
                End If
                Call PH_PY122_AddMatrixRow
        End Select
    End If
    Call oForm.Freeze(False)
    Exit Sub
Raise_FormMenuEvent_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_FormMenuEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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

Sub PH_PY122_AddMatrixRow()
    Dim oRow    As Long
    
    On Error GoTo PH_PY122_AddMatrixRow_Error
    
    Call oForm.Freeze(True)
    
    '//[Mat1]
    oMat1.FlushToDataSource
    oRow = oMat1.VisualRowCount

    If oMat1.VisualRowCount > 0 Then
        If Trim(oDS_PH_PY122B.GetValue("U_MSTCOD", oRow - 1)) <> "" Then
            If oDS_PH_PY122B.Size <= oMat1.VisualRowCount Then
                oDS_PH_PY122B.InsertRecord (oRow)
            End If
            oDS_PH_PY122B.Offset = oRow
            oDS_PH_PY122B.setValue "U_LineNum", oRow, oRow + 1
            oDS_PH_PY122B.setValue "U_MSTCOD", oRow, ""
            oDS_PH_PY122B.setValue "U_MSTNAM", oRow, ""
            oDS_PH_PY122B.setValue "U_TeamCode", oRow, ""
            oDS_PH_PY122B.setValue "U_TeamName", oRow, ""
            oDS_PH_PY122B.setValue "U_RspCode", oRow, ""
            oDS_PH_PY122B.setValue "U_RspName", oRow, ""
            oDS_PH_PY122B.setValue "U_ATeamCod", oRow, ""
            oDS_PH_PY122B.setValue "U_ATeamNam", oRow, ""
            oDS_PH_PY122B.setValue "U_ARspCode", oRow, ""
            oDS_PH_PY122B.setValue "U_ARspName", oRow, ""
            oMat1.LoadFromDataSource
        Else
            oDS_PH_PY122B.Offset = oRow - 1
            
            oDS_PH_PY122B.setValue "U_LineNum", oRow - 1, oRow
            oDS_PH_PY122B.setValue "U_MSTCOD", oRow - 1, ""
            oDS_PH_PY122B.setValue "U_MSTNAM", oRow - 1, ""
            oDS_PH_PY122B.setValue "U_TeamCode", oRow - 1, ""
            oDS_PH_PY122B.setValue "U_TeamName", oRow - 1, ""
            oDS_PH_PY122B.setValue "U_RspCode", oRow - 1, ""
            oDS_PH_PY122B.setValue "U_RspName", oRow - 1, ""
            oDS_PH_PY122B.setValue "U_ATeamCod", oRow - 1, ""
            oDS_PH_PY122B.setValue "U_ATeamNam", oRow - 1, ""
            oDS_PH_PY122B.setValue "U_ARspCode", oRow - 1, ""
            oDS_PH_PY122B.setValue "U_ARspName", oRow - 1, ""
            
            
            oMat1.LoadFromDataSource
        End If
    ElseIf oMat1.VisualRowCount = 0 Then
        oDS_PH_PY122B.Offset = oRow
            oDS_PH_PY122B.setValue "U_LineNum", oRow, oRow + 1
            oDS_PH_PY122B.setValue "U_MSTCOD", oRow, ""
            oDS_PH_PY122B.setValue "U_MSTNAM", oRow, ""
            oDS_PH_PY122B.setValue "U_TeamCode", oRow, ""
            oDS_PH_PY122B.setValue "U_TeamName", oRow, ""
            oDS_PH_PY122B.setValue "U_RspCode", oRow, ""
            oDS_PH_PY122B.setValue "U_RspName", oRow, ""
            oDS_PH_PY122B.setValue "U_ATeamCod", oRow, ""
            oDS_PH_PY122B.setValue "U_ATeamNam", oRow, ""
            oDS_PH_PY122B.setValue "U_ARspCode", oRow, ""
            oDS_PH_PY122B.setValue "U_ARspName", oRow, ""
        oMat1.LoadFromDataSource
    End If
    
    Call oForm.Freeze(False)
    Exit Sub
PH_PY122_AddMatrixRow_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY122_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY122_FormClear()
On Error GoTo PH_PY122_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY122'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PH_PY122_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY122_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY122_DataValidCheck() As Boolean
On Error GoTo PH_PY122_DataValidCheck_Error
    PH_PY122_DataValidCheck = False
    Dim i       As Long
    Dim sQry    As String
    Dim tCode   As String
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    If Trim(oDS_PH_PY122A.GetValue("U_CLTCOD", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "사업장은 필수입니다.", bmt_Short, True
        oForm.Items("CLTCOD").CLICK ct_Regular
        PH_PY122_DataValidCheck = False
        Exit Function
    End If

    
    '//코드 생성
    tCode = Trim(oDS_PH_PY122A.GetValue("U_CLTCOD", 0))
    
    '//코드 중복 체크
    If oForm.Mode = fm_ADD_MODE Then
        sQry = "SELECT CODE FROM [@PH_PY122A] WHERE CODE = '" & tCode & "'"
        oRecordSet.DoQuery sQry
        If oRecordSet.RecordCount > 0 Then
            Sbo_Application.SetStatusBarMessage "코드가 존재합니다", bmt_Short, True
            PH_PY122_DataValidCheck = False
            Exit Function
        Else
            Call oDS_PH_PY122A.setValue("Code", 0, tCode)
            Call oDS_PH_PY122A.setValue("Name", 0, tCode)
        End If
    End If
    
'    If oForm.Mode = fm_EDIT_MODE Then
'        Call oDS_PH_PY122A.setValue("Code", 0, tCode)
'    End If
    
    '//매트릭스 체크
    If oMat1.VisualRowCount > 1 Then
        For i = 1 To oMat1.VisualRowCount - 1
            If oMat1.Columns("MSTCOD").Cells(i).Specific.VALUE = "" Then
                Sbo_Application.SetStatusBarMessage "사번은 필수입니다.", bmt_Short, True
                oMat1.Columns("MSTCOD").Cells(i).CLICK ct_Regular
                PH_PY122_DataValidCheck = False
                Exit Function
            End If
            
            If oMat1.Columns("TeamCode").Cells(i).Specific.VALUE = "" Then
                Sbo_Application.SetStatusBarMessage "부서코드는 필수입니다.", bmt_Short, True
                oMat1.Columns("TeamCode").Cells(i).CLICK ct_Regular
                PH_PY122_DataValidCheck = False
                Exit Function
            End If
            
            If oMat1.Columns("ATeamCod").Cells(i).Specific.VALUE = "" Then
                Sbo_Application.SetStatusBarMessage "변경부서코드는 필수입니다.", bmt_Short, True
                oMat1.Columns("ATeamCod").Cells(i).CLICK ct_Regular
                PH_PY122_DataValidCheck = False
                Exit Function
            End If
            
            If oMat1.Columns("ARspCode").Cells(i).Specific.VALUE = "" Then
                Sbo_Application.SetStatusBarMessage "변경담당코드는 필수입니다.", bmt_Short, True
                oMat1.Columns("ARspCode").Cells(i).CLICK ct_Regular
                PH_PY122_DataValidCheck = False
                Exit Function
            End If
            
        Next i
    Else
        Sbo_Application.SetStatusBarMessage "라인 데이터가 없습니다.", bmt_Short, True
        PH_PY122_DataValidCheck = False
        Exit Function
    End If
    
    
    oMat1.FlushToDataSource
    
    
    '// Matrix 마지막 행 삭제(DB 저장시)
    If oDS_PH_PY122B.Size > 1 Then oDS_PH_PY122B.RemoveRecord (oDS_PH_PY122B.Size - 1)
    
    oMat1.LoadFromDataSource
    
    PH_PY122_DataValidCheck = True
    
    Exit Function
    
    Set oRecordSet = Nothing
    
    
PH_PY122_DataValidCheck_Error:

    Set oRecordSet = Nothing
    PH_PY122_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY122_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY122_MTX01()

    '//메트릭스에 데이터 로드
    
    Dim i       As Long
    Dim sQry    As String
    
    Dim Param01 As String
    Dim Param02 As String
    Dim Param03 As String
    Dim Param04 As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY122_MTX01_Error

    Call oForm.Freeze(True)
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    Param01 = oForm.Items("Param01").Specific.VALUE
    Param02 = oForm.Items("Param01").Specific.VALUE
    Param03 = oForm.Items("Param01").Specific.VALUE
    Param04 = oForm.Items("Param01").Specific.VALUE
    
    sQry = "SELECT 10"
    Call oRecordSet.DoQuery(sQry)
    
    oMat1.Clear
    oMat1.FlushToDataSource
    oMat1.LoadFromDataSource
    
    If (oRecordSet.RecordCount = 0) Then
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
        GoTo PH_PY122_MTX01_Exit
    End If
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet.RecordCount, False)
    
    For i = 0 To oRecordSet.RecordCount - 1
        If i <> 0 Then
            oDS_PH_PY122B.InsertRecord (i)
        End If
        oDS_PH_PY122B.Offset = i
        Call oDS_PH_PY122B.setValue("U_COL01", i, oRecordSet.Fields(0).VALUE)
        Call oDS_PH_PY122B.setValue("U_COL02", i, oRecordSet.Fields(1).VALUE)
        oRecordSet.MoveNext
        ProgressBar01.VALUE = ProgressBar01.VALUE + 1
        ProgressBar01.Text = ProgressBar01.VALUE & "/" & oRecordSet.RecordCount & "건 조회중...!"
    Next
    oMat1.LoadFromDataSource
    oMat1.AutoResizeColumns
    oForm.Update
    
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Sub
PH_PY122_MTX01_Exit:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Exit Sub
PH_PY122_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY122_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY122_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY122_Validate_Error
    PH_PY122_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY122A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", bmt_Short, True
        PH_PY122_Validate = False
        GoTo PH_PY122_Validate_Exit
    End If
'
    If ValidateType = "수정" Then

    ElseIf ValidateType = "행삭제" Then

    ElseIf ValidateType = "취소" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY122_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY122_Validate_Error:
    PH_PY122_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY122_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


