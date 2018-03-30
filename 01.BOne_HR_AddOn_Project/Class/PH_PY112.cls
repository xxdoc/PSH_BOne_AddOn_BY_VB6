VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY112"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY112.cls
'//  Module         : 인사관리 > 급여관리
'//  Desc           : 급상여자료관리
'//********************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form

Private oDS_PH_PY112A As SAPbouiCOM.DBDataSource

Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long

Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY112.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY112_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY112"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_FIND_MODE
    oForm.DataBrowser.BrowseBy = "Code"
    oForm.PaneLevel = 1
        
    oForm.Freeze True
    Call PH_PY112_CreateItems
    Call PH_PY112_EnableMenus
    Call PH_PY112_SetDocument(oFromDocEntry01)
'    Call PH_PY112_FormResize
    
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

Private Function PH_PY112_CreateItems() As Boolean

    Dim sQry        As String
    Dim i           As Long
    
    Dim oCheck      As SAPbouiCOM.CheckBox
    Dim oEdit       As SAPbouiCOM.EditText
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns    As SAPbouiCOM.Columns
    Dim optBtn      As SAPbouiCOM.OptionBtn
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY112_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Set oDS_PH_PY112A = oForm.DataSources.DBDataSources("@PH_PY112A")   '//헤더

    '//----------------------------------------------------------------------------------------------
    '// 아이템 설정
    '//----------------------------------------------------------------------------------------------
    '// 사업장
    Set oCombo = oForm.Items("CLTCOD").Specific
'    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
'    oRecordSet.DoQuery sQry
'    Do Until oRecordSet.EOF
'        oCombo.ValidValues.Add Trim$(oRecordSet.Fields(0).Value), Trim$(oRecordSet.Fields(1).Value)
'        oRecordSet.MoveNext
'    Loop
    oForm.Items("CLTCOD").DisplayDesc = True
    
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
    
    '//부서
    oForm.Items("TeamCode").DisplayDesc = True

    '//담당
    oForm.Items("RspCode").DisplayDesc = True
    
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
    
    '// 1.급여형태
    Set oCombo = oForm.Items("PAYTYP").Specific
    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P132' ORDER BY CAST(U_Code AS NUMERIC) "
    Call SetReDataCombo(oForm, sQry, oCombo)
    oForm.Items("PAYTYP").DisplayDesc = True
    
    '// 2.직급형태
    Set oCombo = oForm.Items("JIGCOD").Specific
    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P129' ORDER BY U_Code "
    Call SetReDataCombo(oForm, sQry, oCombo)
    oForm.Items("JIGCOD").DisplayDesc = True
    
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
    
PH_PY112_CreateItems_Error:
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY112_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY112_EnableMenus()
    
    On Error GoTo PH_PY112_EnableMenus_Error

    Call oForm.EnableMenu("1283", True)     '//제거
    Call oForm.EnableMenu("1284", False)    '//취소
    Call oForm.EnableMenu("1293", True)     '//행삭제
    
    Exit Sub
    
PH_PY112_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY112_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY112_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY112_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY112_FormItemEnabled
        Call PH_PY112_AddMatrixRow
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY112_FormItemEnabled
        oForm.Items("Code").Specific.VALUE = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY112_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY112_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY112_FormItemEnabled()
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim i           As Long
    Dim sQry        As String
    On Error GoTo PH_PY112_FormItemEnabled_Error
    
    
    Call oForm.Freeze(True)
    If (oForm.Mode = fm_ADD_MODE) Then
        oForm.Items("Code").Enabled = True
    
        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", False)    '//문서추가
        
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
        
        If oDS_PH_PY112A.GetValue("U_CLTCOD", 0) <> "" Then
            sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
            sQry = sQry & " WHERE Code = '1' AND U_Char2 = '" & Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) & "'"
            sQry = sQry & " ORDER BY U_Code"
            Call SetReDataCombo(oForm, sQry, oCombo)
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
        
        If oForm.Items.Item("CLTCOD").Specific.VALUE <> "" Then
            sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
            sQry = sQry & " WHERE Code = '2' AND U_Char2 = '" & Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) & "'"
            sQry = sQry & " Order By U_Code"
            Call SetReDataCombo(oForm, sQry, oCombo)
        End If
        
        oForm.Mode = fm_FIND_MODE
        Call PH_PY112_ItemNameSetting
    ElseIf (oForm.Mode = fm_FIND_MODE) Then
        oForm.Items("Code").Enabled = True
        
        Call oForm.EnableMenu("1281", True)    '//문서찾기
        Call oForm.EnableMenu("1282", False)     '//문서추가
        
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
        
        Call PH_PY112_ItemNameSetting
    ElseIf (oForm.Mode = fm_OK_MODE) Then
        oForm.Items("Code").Enabled = False
        
        
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD", False)
        
        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", False)     '//문서추가
        
        Call PH_PY112_ItemNameSetting
    End If
    Call oForm.Freeze(False)
    Exit Sub
    
PH_PY112_FormItemEnabled_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY112_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
                If pval.ItemUID = "1" Then
                    If oForm.Mode = fm_ADD_MODE Then
                        If PH_PY112_DataValidCheck = False Then
                            BubbleEvent = False
                        End If
                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
                        If PH_PY112_DataValidCheck = False Then
                            BubbleEvent = False
                        End If
                    ElseIf oForm.Mode = fm_OK_MODE Then
                    End If
                End If
                If pval.ItemUID = "Btn01" Then
                    oForm.Items("HOBONG").CLICK ct_Regular
                    Sbo_Application.ActivateMenuItem ("7425")
                    BubbleEvent = False
                End If
            ElseIf pval.BeforeAction = False Then
                If Mid(pval.ItemUID, 1, 6) = "Folder" Then
                    oForm.PaneLevel = Right(pval.ItemUID, 1)
                End If
                If pval.ItemUID = "1" Then
                    If oForm.Mode = fm_ADD_MODE Then
                        If pval.ActionSuccess = True Then
'                            Call PH_PY112_FormItemEnabled
'                            Call PH_PY112_AddMatrixRow
                        End If
                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
                        If pval.ActionSuccess = True Then
'                            Call PH_PY112_FormItemEnabled
'                            Call PH_PY112_AddMatrixRow
                        End If
                    ElseIf oForm.Mode = fm_OK_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY112_FormItemEnabled
                        End If
                    ElseIf oForm.Mode = fm_FIND_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY112_FormItemEnabled
                        End If
                    End If
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
                    '//사업장(헤더)
                    If pval.ItemUID = "CLTCOD" Then
                        
                        '//기본사항 - 부서 (사업장에 따른 부서변경)
                        Set oCombo = oForm.Items("TeamCode").Specific
    
                        If oCombo.ValidValues.Count > 0 Then
                            For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                                oCombo.ValidValues.Remove i, psk_Index
                            Next i
                            oCombo.ValidValues.Add "", ""
                            oCombo.Select 0, psk_Index
                        End If
                        
                        sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
                        sQry = sQry & " WHERE Code = '1' AND U_Char2 = '" & Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) & "'"
                        sQry = sQry & " ORDER BY U_Code"
                        Call SetReDataCombo(oForm, sQry, oCombo)
                        
                        oForm.Items("TeamCode").DisplayDesc = True
                        
                        '//담당 (사업장에 따른 담당변경)
                    
                        Set oCombo = oForm.Items("RspCode").Specific
    
                        If oCombo.ValidValues.Count > 0 Then
                            For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                                oCombo.ValidValues.Remove i, psk_Index
                            Next i
                            oCombo.ValidValues.Add "", ""
                            oCombo.Select 0, psk_Index
                        End If
                        
                        sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
                        sQry = sQry & " WHERE Code = '2' AND U_Char2 = '" & Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) & "'"
                        sQry = sQry & " Order By U_Code"
                        Call SetReDataCombo(oForm, sQry, oCombo)
                        
                        oForm.Items("RspCode").DisplayDesc = True
                        
                        Call PH_PY112_ItemNameSetting
                    End If
                End If
            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_CLICK: '//6
            If pval.BeforeAction = True Then
                Select Case pval.ItemUID
                Case "Mat1", "Grid1"
                    If pval.Row > 0 Then
                        '//Call oMat1.SelectRow(pval.Row, True, False)

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
        Case et_MATRIX_LOAD: '//11
'            If pval.BeforeAction = True Then
'            ElseIf pval.BeforeAction = False Then
'                oMat1.LoadFromDataSource
'
'                Call PH_PY112_FormItemEnabled
'                Call PH_PY112_AddMatrixRow
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
'                Set oDS_PH_PY112A = Nothing
'                Set oDS_PH_PY112B = Nothing
'
'                Set oMat1 = Nothing
'                Set oGrid1 = Nothing
                
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
                Call AuthorityCheck(oForm, "CLTCOD", "@PH_PY112A", "Code")      '//접속자 권한에 따른 사업장 보기
        End Select
    ElseIf (pval.BeforeAction = False) Then
        Select Case pval.MenuUID
            Case "1283":
                oForm.Mode = fm_ADD_MODE
                Call PH_PY112_FormItemEnabled
                Call PH_PY112_AddMatrixRow
            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//문서찾기
                Call PH_PY112_FormItemEnabled
                Call PH_PY112_AddMatrixRow
                oForm.Items("Code").CLICK ct_Regular
            Case "1282":    '//문서추가
                Call PH_PY112_FormItemEnabled
                Call PH_PY112_AddMatrixRow
            Case "1288", "1289", "1290", "1291":
                Call PH_PY112_FormItemEnabled
            Case "1293" '// 행삭제
'                '// [MAT1 용]
'                 If oMat1.RowCount <> oMat1.VisualRowCount Then
'                    oMat1.FlushToDataSource
'
'                    While (i <= oDS_PH_PY112B.Size - 1)
'                        If oDS_PH_PY112B.GetValue("U_FILD01", i) = "" Then
'                            oDS_PH_PY112B.RemoveRecord (i)
'                            i = 0
'                        Else
'                            i = i + 1
'                        End If
'                    Wend
'
'                    For i = 0 To oDS_PH_PY112B.Size
'                        Call oDS_PH_PY112B.setValue("U_LineNum", i, i + 1)
'                    Next i
'
'                    oMat1.LoadFromDataSource
'                End If
'                Call PH_PY112_AddMatrixRow
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
                '//부서
                Set oCombo = oForm.Items("TeamCode").Specific
                If oCombo.ValidValues.Count > 0 Then
                    For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                        oCombo.ValidValues.Remove i, psk_Index
                    Next i
                End If
                
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
                sQry = sQry & " WHERE Code = '1' AND U_Char2 = '" & Trim(oDS_PH_PY112A.GetValue("U_CLTCOD", 0)) & "'"
                sQry = sQry & " ORDER BY U_Code"

                Call SetReDataCombo(oForm, sQry, oCombo)

                oForm.Items("TeamCode").DisplayDesc = True
                
        
                '//담당
                Set oCombo = oForm.Items("RspCode").Specific
                If oCombo.ValidValues.Count > 0 Then
                    For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                        oCombo.ValidValues.Remove i, psk_Index
                    Next i
                End If
                
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
                sQry = sQry & " WHERE Code = '2' AND U_Char2 = '" & Trim(oDS_PH_PY112A.GetValue("U_CLTCOD", 0)) & "'"
                sQry = sQry & " ORDER BY U_Code"

                Call SetReDataCombo(oForm, sQry, oCombo)

                oForm.Items("RspCode").DisplayDesc = True
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

Sub PH_PY112_AddMatrixRow()
    Dim oRow    As Long
    
    On Error GoTo PH_PY112_AddMatrixRow_Error
    
    Call oForm.Freeze(True)
    
'    '//[Mat1 용]
'    oMat1.FlushToDataSource
'    oRow = oMat1.VisualRowCount
'
'    If oMat1.VisualRowCount > 0 Then
'        If Trim(oDS_PH_PY112B.GetValue("U_FILD01", oRow - 1)) <> "" Then
'            If oDS_PH_PY112B.Size <= oMat1.VisualRowCount Then
'                oDS_PH_PY112B.InsertRecord (oRow)
'            End If
'            oDS_PH_PY112B.Offset = oRow
'            oDS_PH_PY112B.setValue "U_LineNum", oRow, oRow + 1
'            oDS_PH_PY112B.setValue "U_FILD01", oRow, ""
'            oDS_PH_PY112B.setValue "U_FILD02", oRow, ""
'            oDS_PH_PY112B.setValue "U_FILD03", oRow, 0
'            oMat1.LoadFromDataSource
'        Else
'            oDS_PH_PY112B.Offset = oRow - 1
'            oDS_PH_PY112B.setValue "U_LineNum", oRow - 1, oRow
'            oDS_PH_PY112B.setValue "U_FILD01", oRow - 1, ""
'            oDS_PH_PY112B.setValue "U_FILD02", oRow - 1, ""
'            oDS_PH_PY112B.setValue "U_FILD03", oRow - 1, 0
'            oMat1.LoadFromDataSource
'        End If
'    ElseIf oMat1.VisualRowCount = 0 Then
'        oDS_PH_PY112B.Offset = oRow
'        oDS_PH_PY112B.setValue "U_LineNum", oRow, oRow + 1
'        oDS_PH_PY112B.setValue "U_FILD01", oRow, ""
'        oDS_PH_PY112B.setValue "U_FILD02", oRow, ""
'        oDS_PH_PY112B.setValue "U_FILD03", oRow, 0
'        oMat1.LoadFromDataSource
'    End If
    
    Call oForm.Freeze(False)
    Exit Sub
PH_PY112_AddMatrixRow_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY112_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY112_FormClear()
On Error GoTo PH_PY112_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY112'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PH_PY112_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY112_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY112_DataValidCheck() As Boolean
On Error GoTo PH_PY112_DataValidCheck_Error
    PH_PY112_DataValidCheck = False
    Dim i       As Long
    Dim sQry    As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '//----------------------------------------------------------------------------------
    '//필수 체크
    '//----------------------------------------------------------------------------------
'    If Trim(oDS_PH_PY112A.GetValue("Code", 0)) = "" Then
'        Sbo_Application.SetStatusBarMessage "사원번호는 필수입니다.", bmt_Short, True
'        oForm.Items("Code").CLICK ct_Regular
'        Exit Function
'    End If
'
'    oMat1.FlushToDataSource
'    '// Matrix 마지막 행 삭제(DB 저장시)
'    If oDS_PH_PY112B.Size > 1 Then oDS_PH_PY112B.RemoveRecord (oDS_PH_PY112B.Size - 1)
'    oMat1.LoadFromDataSource
    
    PH_PY112_DataValidCheck = True
    
    Exit Function
    
    Set oRecordSet = Nothing
    
    
PH_PY112_DataValidCheck_Error:
    Set oRecordSet = Nothing
    PH_PY112_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY112_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Function PH_PY112_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY112_Validate_Error
    PH_PY112_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY112A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", bmt_Short, True
        PH_PY112_Validate = False
        GoTo PH_PY112_Validate_Exit
    End If
'
    If ValidateType = "수정" Then

    ElseIf ValidateType = "행삭제" Then

    ElseIf ValidateType = "취소" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY112_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY112_Validate_Error:
    PH_PY112_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY112_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function



Private Sub PH_PY112_Print_Report01()

    Dim DocNum          As String
    Dim ErrNum          As Integer
    Dim WinTitle        As String
    Dim ReportName      As String
    Dim sQry            As String
    
    Dim BPLID           As String
    Dim ItmBsort        As String
    Dim DocDate         As String
    
    Dim oRecordSet      As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY112_Print_Report01_Error
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '/ ODBC 연결 체크
    If ConnectODBC = False Then
        GoTo PH_PY112_Print_Report01_Error
    End If
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/

    WinTitle = "[S142] 발주서"
    ReportName = "S142_1.rpt"
    sQry = "EXEC PH_PY112_1 '" & oForm.Items("8").Specific.VALUE & "'"
    
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
    
PH_PY112_Print_Report01_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY112_Print_Report01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Sub PH_PY112_ItemNameSetting()
On Error GoTo PH_PY112_ItemNameSetting_Error
    Dim i           As Long
    Dim sQry        As String
    Dim oStatic     As SAPbouiCOM.StaticText
    Dim YM          As String
    Dim CLTCOD      As String
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY112_ItemNameSetting_Error
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    YM = Trim(oDS_PH_PY112A.GetValue("U_YM", 0))
    CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE)
    oForm.PaneLevel = 1
    
    '//지급 항목
    sQry = "SELECT T0.U_CSUNAM"
    sQry = sQry & " FROM [@PH_PY102B] T0 INNER JOIN [@PH_PY102A] T1 ON T0.Code = T1.Code"
    sQry = sQry & " WHERE T1.U_YM = '" & YM & "' OR"
    sQry = sQry & " (T1.U_YM <> '" & YM & "' AND T1.U_YM = (SELECT MAX(U_YM) FROM [@PH_PY102A] WHERE U_CLTCOD = '" & CLTCOD & "' ))"
    sQry = sQry & " AND T1.U_CLTCOD = '" & CLTCOD & "'"
    oRecordSet.DoQuery sQry
    
        For i = 1 To 36
            If i <= oRecordSet.RecordCount Then
                Set oStatic = oForm.Items("sCSUD" & Format(i, "00")).Specific
                oStatic.Caption = Format(i, "00") & ") " & oRecordSet.Fields(0).VALUE
                oForm.Items("sCSUD" & Format(i, "00")).Visible = True
                oForm.Items("CSUD" & Format(i, "00")).Visible = True
                oRecordSet.MoveNext
            Else
                oForm.Items("sCSUD" & Format(i, "00")).Visible = False
                oForm.Items("CSUD" & Format(i, "00")).Visible = False
            End If
        Next i
    
    '//공제 항목
    sQry = "SELECT T0.U_CSUNAM"
    sQry = sQry & " FROM [@PH_PY103B] T0 INNER JOIN [@PH_PY103A] T1 ON T0.Code = T1.Code"
    sQry = sQry & " WHERE T1.U_YM = '" & YM & "' OR "
    sQry = sQry & " (T1.U_YM <> '" & YM & "' AND T1.U_YM = (SELECT MAX(U_YM) FROM [@PH_PY103A] WHERE U_CLTCOD = '" & CLTCOD & "' ))"
    sQry = sQry & " AND T1.U_CLTCOD = '" & CLTCOD & "'"
    oRecordSet.DoQuery sQry
    
    For i = 1 To 36
        If i <= oRecordSet.RecordCount Then
            Set oStatic = oForm.Items("sGONG" & Format(i, "00")).Specific
            oStatic.Caption = Format(i, "00") & ") " & oRecordSet.Fields(0).VALUE
            oForm.Items("sGONG" & Format(i, "00")).Visible = True
            oForm.Items("GONG" & Format(i, "00")).Visible = True
            oRecordSet.MoveNext
        Else
            oForm.Items("sGONG" & Format(i, "00")).Visible = False
            oForm.Items("GONG" & Format(i, "00")).Visible = False
        End If
    Next i

    oForm.PaneLevel = 2
    '//비과세 항목
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P218' AND U_UseYN= 'Y'"
    oRecordSet.DoQuery sQry
'        sQry = "SELECT T0.U_CSUNAM"
'    sQry = sQry & " FROM [@PH_PY103B] T0 INNER JOIN [@PH_PY103A] T1 ON T0.Code = T1.Code"
'    sQry = sQry & " WHERE T1.U_YM = '" & YM & "' OR (T1.U_YM <> '" & YM & "' AND T1.U_YM = (SELECT MAX(U_YM) FROM [@PH_PY103A]))"
'    oRecordSet.DoQuery sQry
    
    For i = 1 To 36
        If i <= oRecordSet.RecordCount Then
            Set oStatic = oForm.Items("sBTX" & Format(i, "00")).Specific
            oStatic.Caption = Format(i, "00") & ") " & oRecordSet.Fields(0).VALUE
            oForm.Items("sBTX" & Format(i, "00")).Visible = True
            oForm.Items("BTX" & Format(i, "00")).Visible = True
            oRecordSet.MoveNext
        Else
            oForm.Items("sBTX" & Format(i, "00")).Visible = False
            oForm.Items("BTX" & Format(i, "00")).Visible = False
        End If
    Next i
 
    oForm.PaneLevel = 1
    
    Set oRecordSet = Nothing
    Exit Sub
    
PH_PY112_ItemNameSetting_Error:
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "PH_PY112_ItemNameSetting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

