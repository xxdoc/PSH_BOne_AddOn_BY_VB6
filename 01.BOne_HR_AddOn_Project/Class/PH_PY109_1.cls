VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY109_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY109_1.cls
'//  Module         : 인사관리 > 급여관리
'//  Desc           : 급상여변동자료등록
'//********************************************************************************
Option Explicit

Public oFormUniqueID            As String
Public oForm                    As SAPbouiCOM.Form

Public oMat1                    As SAPbouiCOM.Matrix

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
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY109_1.srf")
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
    
    oFormUniqueID = "PH_PY109_1_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY109_1"
    
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
'    oForm.DataBrowser.BrowseBy = "Code"
        
    oForm.Freeze True
    Call PH_PY109_1_CreateItems
    Call PH_PY109_1_EnableMenus
    
'    Call PH_PY109_1_FormResize
    
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

Private Function PH_PY109_1_CreateItems() As Boolean

    Dim sQry        As String
    Dim i           As Long
    
    Dim oCheck      As SAPbouiCOM.CheckBox
    Dim oEdit       As SAPbouiCOM.EditText
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns    As SAPbouiCOM.Columns
    Dim optBtn      As SAPbouiCOM.OptionBtn
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY109_1_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '/Matrix
    Set oDS_PH_PY109Z = oForm.DataSources.DBDataSources("@PS_USERDS01")   '//라인

    Set oMat1 = oForm.Items("Mat1").Specific
    
    
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
    
    '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
    
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
    
    
    oMat1.Columns("PayDud").ValidValues.Add "1", "수당"
    oMat1.Columns("PayDud").ValidValues.Add "2", "공제"
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
    
PH_PY109_1_CreateItems_Error:
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY109_1_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Private Sub PH_PY109_1_EnableMenus()
    
    On Error GoTo PH_PY109_1_EnableMenus_Error

    Call oForm.EnableMenu("1281", False)     '//찾기
    Call oForm.EnableMenu("1283", False)     '//제거
    Call oForm.EnableMenu("1284", False)    '//취소
    Call oForm.EnableMenu("1293", False)     '//행삭제
    
    Exit Sub
    
PH_PY109_1_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY109_1_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub



Sub PH_PY109_1_FormItemEnabled()
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oConditions As SAPbouiCOM.Conditions
    
    On Error GoTo PH_PY109_1_FormItemEnabled_Error
    
    
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
        
        Call oForm.EnableMenu("1281", False)     '//문서찾기
        Call oForm.EnableMenu("1282", False)    '//문서추가
        
        
    End If
    Call oForm.Freeze(False)
    Exit Sub
    
PH_PY109_1_FormItemEnabled_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY109_1_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
                If pval.ItemUID = "Btn_01" Then
                    If oForm.Mode = fm_ADD_MODE Or oForm.Mode = fm_UPDATE_MODE Then
                        If PH_PY109_1_DataValidCheck(pval.ItemUID) = False Then
                            BubbleEvent = False
                        Else
                            Call PH_PY109_1_SaveData
                        End If
                        
                    ElseIf oForm.Mode = fm_FIND_MODE Then
                    End If
                End If
                If pval.ItemUID = "Btn_Ret" Then
                    If PH_PY109_1_DataValidCheck(pval.ItemUID) = False Then
                        BubbleEvent = False
                    Else
                        Call PH_PY109_1_LoadData
                    End If
                End If
            ElseIf pval.BeforeAction = False Then
                
            End If
     
'----------------------------------------------------------
        Case et_KEY_DOWN: '//2
             If pval.BeforeAction = True And pval.ItemUID = "Mat1" And pval.ColUID = "PDCode" And pval.CharPressed = 9 Then
                If oMat1.Columns("PDCode").Cells(pval.Row).Specific.VALUE = "" Then
                    oMat1.Columns("PDCode").Cells(pval.Row).CLICK ct_Regular
                    Sbo_Application.ActivateMenuItem ("7425")
                    BubbleEvent = False
                
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
                    If pval.ItemUID = "Mat1" And pval.ColUID = "PDCode" Then
                        If oMat1.Columns("PDCode").Cells(pval.Row).Specific.VALUE <> "" Then
                            If oMat1.Columns("PayDud").Cells(pval.Row).Specific.VALUE = "1" Then
                                oMat1.Columns("PDName").Cells(pval.Row).Specific.VALUE = MDC_SetMod.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" & oMat1.Columns("PDCode").Cells(pval.Row).Specific.VALUE & "'", " AND Code = 'P134'")
                            Else
                                oMat1.Columns("PDName").Cells(pval.Row).Specific.VALUE = MDC_SetMod.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" & oMat1.Columns("PDCode").Cells(pval.Row).Specific.VALUE & "'", " AND Code = 'P143'")
                            End If
                        End If
                        
                        oMat1.Columns("PDCode").Cells(pval.Row).CLICK ct_Regular
                    End If
                    
                    
                End If
            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_MATRIX_LOAD: '//11
            If pval.BeforeAction = True Then
            ElseIf pval.BeforeAction = False Then
                oMat1.LoadFromDataSource

                Call PH_PY109_1_FormItemEnabled
                
                

                
                
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
                Set oDS_PH_PY109Z = Nothing
        
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
            If pval.BeforeAction = True Then

            ElseIf pval.Before_Action = False Then
                If pval.ItemUID = "Mat1" Then
                    If pval.ColUID = "MSTCOD" Then      '// 프로젝트 코드
                        Call MDC_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY109_1B", "U_MSTCOD,U_MSTNAM,U_DPTCOD,U_DPTNAM", "Mat1", pval.Row)
                        
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
                
        End Select
    ElseIf (pval.BeforeAction = False) Then
        Select Case pval.MenuUID
            Case "1283":
                oForm.Mode = fm_ADD_MODE
                Call PH_PY109_1_FormItemEnabled
                
                
            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//문서찾기
                Call PH_PY109_1_FormItemEnabled
                
                oForm.Items("Code").CLICK ct_Regular
            Case "1282":    '//문서추가
                Call PH_PY109_1_FormItemEnabled
                
            Case "1288", "1289", "1290", "1291":
                Call PH_PY109_1_FormItemEnabled
                
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
    
    Exit Sub
Raise_RightClickEvent_Error:

    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub



Sub PH_PY109_1_FormClear()
On Error GoTo PH_PY109_1_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY109_1'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PH_PY109_1_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY109_1_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY109_1_DataValidCheck(ItemUID As String) As Boolean
On Error GoTo PH_PY109_1_DataValidCheck_Error
    PH_PY109_1_DataValidCheck = False
    Dim i       As Long
    Dim sQry    As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '//----------------------------------------------------------------------------------
    '//필수 체크
    '//----------------------------------------------------------------------------------
    If Trim(oForm.Items("CLTCOD").Specific.VALUE) = "" Then
        Sbo_Application.SetStatusBarMessage "사업장은 필수입니다.", bmt_Short, True
        oForm.Items("CLTCOD").CLICK ct_Regular
        Exit Function
    End If
    
    If Trim(oForm.Items("YM").Specific.VALUE) = "" Then
        Sbo_Application.SetStatusBarMessage "귀속년월은 필수입니다.", bmt_Short, True
        oForm.Items("YM").CLICK ct_Regular
        Exit Function
    End If
    
    If Trim(oForm.Items("JOBTYP").Specific.VALUE) = "" Then
        Sbo_Application.SetStatusBarMessage "지급종류는 필수입니다.", bmt_Short, True
        oForm.Items("JOBTYP").CLICK ct_Regular
        Exit Function
    End If
    
    If Trim(oForm.Items("JOBGBN").Specific.VALUE) = "" Then
        Sbo_Application.SetStatusBarMessage "지급구분은 필수입니다.", bmt_Short, True
        oForm.Items("JOBGBN").CLICK ct_Regular
        Exit Function
    End If
    
    If Trim(oForm.Items("JOBTRG").Specific.VALUE) = "" Then
        Sbo_Application.SetStatusBarMessage "지급대상은 필수입니다.", bmt_Short, True
        oForm.Items("JOBTRG").CLICK ct_Regular
        Exit Function
    End If
    
    
    
        
    oMat1.FlushToDataSource
    
    
    If ItemUID = "Btn_01" Then
        '// Matrix 마지막 행 삭제(DB 저장시)
        If oDS_PH_PY109Z.Size < 1 Then
            Sbo_Application.SetStatusBarMessage "라인이 존재하지 않습니다.", bmt_Short, True
            PH_PY109_1_DataValidCheck = False
            Exit Function
        End If
        
        Dim Code As String
        Dim PDCode As String
        Dim bPDCode As String
        Dim Amt As Double
        
        For i = 0 To oDS_PH_PY109Z.Size - 1 ' oMat1.VisualRowCount - 1
            Code = Trim(oDS_PH_PY109Z.GetValue("U_ColReg01", i))
            PDCode = Trim(oDS_PH_PY109Z.GetValue("U_ColReg03", i))
            bPDCode = Trim(oDS_PH_PY109Z.GetValue("U_ColReg05", i))
            Amt = Trim(oDS_PH_PY109Z.GetValue("U_ColSum01", i))
            
            If Code <> "" Then '//문서번호가 있고
                
                If PDCode <> bPDCode Then '//수당코드가 수정이 되었으면 Update대상
                    If Amt <> 0 Then
                       Call MDC_Com.MDC_GF_Message("변동자료에 금액이있는 항목은 수정할 수 없습니다.", "W")
                       PH_PY109_1_DataValidCheck = False
                        Exit Function
                    End If
                End If
            End If
        Next
    
    End If
    oMat1.LoadFromDataSource
    
    PH_PY109_1_DataValidCheck = True
    
    Exit Function
    
    Set oRecordSet = Nothing
    
    
PH_PY109_1_DataValidCheck_Error:
    Set oRecordSet = Nothing
    PH_PY109_1_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY109_1_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Function PH_PY109_1_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY109_1_Validate_Error
    PH_PY109_1_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY109_1A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", bmt_Short, True
        PH_PY109_1_Validate = False
        GoTo PH_PY109_1_Validate_Exit
    End If
'
    If ValidateType = "수정" Then

    ElseIf ValidateType = "행삭제" Then

    ElseIf ValidateType = "취소" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY109_1_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY109_1_Validate_Error:
    PH_PY109_1_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY109_1_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Function PH_PY109_1_LoadData()
    Dim i           As Long
    Dim sQry        As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    Dim oCombo      As SAPbouiCOM.ComboBox

    Dim CLTCOD As String
    Dim YM As String
    Dim JOBTYP As String
    Dim JOBGBN As String
    Dim JOBTRG As String
    
    On Error GoTo PH_PY109_1_DataLoad_ERROR
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    oForm.Freeze (True)
    
    CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE)
    YM = oForm.Items("YM").Specific.VALUE
    JOBTYP = oForm.Items("JOBTYP").Specific.VALUE
    JOBGBN = oForm.Items("JOBGBN").Specific.VALUE
    JOBTRG = oForm.Items("JOBTRG").Specific.VALUE
        
    '// 수당, 공제 테이블 고정:V, 상여:Y 인 값을 임시테이블에 넣는다
    sQry = "EXEC PH_PY109_1_01 '" & CLTCOD & "' , '" & YM & "' , '" & JOBTYP & "' , '" & JOBGBN & "' , '" & JOBTRG & "'"
    oRecordSet.DoQuery sQry

    oMat1.Clear
    oMat1.FlushToDataSource
    oMat1.LoadFromDataSource
    
    If (oRecordSet.RecordCount = 0) Then
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "W")
        Exit Function
    Else
        For i = 0 To oRecordSet.RecordCount - 1
            If i <> 0 Then
                oDS_PH_PY109Z.InsertRecord (i)
            End If
            oDS_PH_PY109Z.Offset = i
            'Call oDS_PH_PY109_1Z.setValue("U_LineNum", i, i + 1)
            Call oDS_PH_PY109Z.setValue("U_ColReg01", i, oRecordSet.Fields("Code").VALUE)
            Call oDS_PH_PY109Z.setValue("U_ColNum01", i, oRecordSet.Fields("Sequence").VALUE)
            Call oDS_PH_PY109Z.setValue("U_ColReg02", i, oRecordSet.Fields("PayDud").VALUE)
            Call oDS_PH_PY109Z.setValue("U_ColReg03", i, oRecordSet.Fields("PDCode").VALUE)
            Call oDS_PH_PY109Z.setValue("U_ColReg04", i, oRecordSet.Fields("PDName").VALUE)
            Call oDS_PH_PY109Z.setValue("U_ColSum01", i, oRecordSet.Fields("Amt").VALUE)
            Call oDS_PH_PY109Z.setValue("U_ColReg05", i, oRecordSet.Fields("PDCode").VALUE)
            Call oDS_PH_PY109Z.setValue("U_ColReg06", i, oRecordSet.Fields("PDName").VALUE)
            oRecordSet.MoveNext
'            ProgressBar01.VALUE = ProgressBar01.VALUE + 1
'            ProgressBar01.Text = ProgressBar01.VALUE & "/" & RecordSet01.RecordCount & "건 조회중...!"
        Next
        oMat1.LoadFromDataSource
        oMat1.AutoResizeColumns
        oForm.Update
    
    End If
    oForm.Freeze (False)
    Exit Function
    
PH_PY109_1_DataLoad_ERROR:
    oForm.Freeze (False)
    Sbo_Application.SetStatusBarMessage "PH_PY109_1_DataLoad_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Function PH_PY109_1_SaveData()
    Dim i           As Long
    Dim sQry        As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    Dim oCombo      As SAPbouiCOM.ComboBox

    Dim Code As String
    Dim Sequence As Integer
    
    Dim PDCode As String
    Dim PDName As String
    Dim bPDCode As String
    Dim bPDName As String
    Dim Amt As Double
    
    On Error GoTo PH_PY109_1_SaveData_ERROR
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    oForm.Freeze (True)
    
            
    oMat1.FlushToDataSource
    
    For i = 0 To oDS_PH_PY109Z.Size - 1
        Code = Trim(oDS_PH_PY109Z.GetValue("U_ColReg01", i))
        Sequence = Trim(oDS_PH_PY109Z.GetValue("U_ColNum01", i))
        PDCode = Trim(oDS_PH_PY109Z.GetValue("U_ColReg03", i))
        PDName = Trim(oDS_PH_PY109Z.GetValue("U_ColReg04", i))
        bPDCode = Trim(oDS_PH_PY109Z.GetValue("U_ColReg05", i))
        bPDName = Trim(oDS_PH_PY109Z.GetValue("U_ColReg06", i))
        Amt = Trim(oDS_PH_PY109Z.GetValue("U_ColSum01", i))
        
        If Code <> "" Then '//문서번호가 있고
            
            If PDCode <> bPDCode Then '//수당코드가 수정이 되었으면 Update대상
                If Amt = 0 Then
                    sQry = "Update [@PH_PY109Z] Set U_PDCode = '" & PDCode & "' , U_PDName = '" & PDName & "'"
                    sQry = sQry + " Where Code = '" & Code & "' And U_Sequence = " & Sequence & ""
                    sQry = sQry + " And U_PDCode = '" & bPDCode & "' And U_PDName = '" & bPDName & "'"
                    
                    oRecordSet.DoQuery sQry
                    Call MDC_Com.MDC_GF_Message("저장되었습니다. 급여변동자료 등록에서 확인바랍니다.", "W")
                End If
            End If
        End If
    Next
    Set oRecordSet = Nothing
    oForm.Freeze (False)
    Exit Function
    
PH_PY109_1_SaveData_ERROR:
    oForm.Freeze (False)
    Sbo_Application.SetStatusBarMessage "PH_PY109_1_SaveData_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function
