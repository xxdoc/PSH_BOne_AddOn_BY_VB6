VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY309"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY309.cls
'//  Module         : 인사관리 > 기타 > 대부금관리
'//  Desc           : 대부금등록
'//********************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form

Public oMat1 As SAPbouiCOM.Matrix

Private oDS_PH_PY309A As SAPbouiCOM.DBDataSource
Private oDS_PH_PY309B As SAPbouiCOM.DBDataSource

Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long

Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY309.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY309_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY309"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
    oForm.DataBrowser.BrowseBy = "DocEntry"
        
    oForm.Freeze True
    Call PH_PY309_CreateItems
    Call PH_PY309_EnableMenus
    Call PH_PY309_SetDocument(oFromDocEntry01)
'    Call PH_PY309_FormResize
    
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

Private Function PH_PY309_CreateItems() As Boolean

    Dim sQry        As String
    Dim i           As Long
    
    Dim oEdit       As SAPbouiCOM.EditText
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns    As SAPbouiCOM.Columns
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY309_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Set oDS_PH_PY309A = oForm.DataSources.DBDataSources("@PH_PY309A")
    Set oDS_PH_PY309B = oForm.DataSources.DBDataSources("@PH_PY309B")

    
    Set oMat1 = oForm.Items("Mat01").Specific

    oMat1.SelectionMode = ms_Auto
    oMat1.AutoResizeColumns


    '//----------------------------------------------------------------------------------------------
    '// 기본사항
    '//----------------------------------------------------------------------------------------------
    
    '사업장
    Set oCombo = oForm.Items("CLTCOD").Specific
'    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
'    Call SetReDataCombo(oForm, sQry, oCombo)
    oForm.Items("CLTCOD").DisplayDesc = True

'    '분기
'    Set oCombo = oForm.Items("Quarter").Specific
'    oCombo.ValidValues.Add "", ""
'    oCombo.ValidValues.Add "01", "1/4 혹은 1학기"
'    oCombo.ValidValues.Add "02", "2/4"
'    oCombo.ValidValues.Add "03", "3/4 혹은 2학기"
'    oCombo.ValidValues.Add "04", "4/4"
'    oCombo.Select 0, psk_Index
'    oForm.Items("Quarter").DisplayDesc = True
'
'    '매트릭스-성별
'    Set oColumn = oMat1.Columns("Sex")
'    oColumn.ValidValues.Add "", ""
'    oColumn.ValidValues.Add "01", "남자"
'    oColumn.ValidValues.Add "02", "여자"
'    oColumn.DisplayDesc = True
'
'    '매트릭스-학교
'    Set oColumn = oMat1.Columns("SchCls")
'    oColumn.ValidValues.Add "", ""
'    sQry = "            SELECT      T1.U_Code,"
'    sQry = sQry & "                 T1.U_CodeNm"
'    sQry = sQry & "  FROM       [@PS_HR200H] AS T0"
'    sQry = sQry & "                 INNER JOIN"
'    sQry = sQry & "                 [@PS_HR200L] AS T1"
'    sQry = sQry & "                     ON T0.Code = T1.Code"
'    sQry = sQry & "  WHERE      T0.Code = 'P222'"
'    sQry = sQry & "                 AND T1.U_UseYN = 'Y'"
'    sQry = sQry & "  ORDER BY  T1.U_Seq"
'
'    Call MDC_SetMod.GP_MatrixSetMatComboList(oColumn, sQry, False, False)
'
''    oColumn.ValidValues.Add "01", "고등학교"
''    oColumn.ValidValues.Add "02", "전문대학"
''    oColumn.ValidValues.Add "03", "대학교"
'    oColumn.DisplayDesc = True
'
'    '매트릭스-학년
'    Set oColumn = oMat1.Columns("Grade")
'    oColumn.ValidValues.Add "", ""
'    oColumn.ValidValues.Add "01", "1학년"
'    oColumn.ValidValues.Add "02", "2학년"
'    oColumn.ValidValues.Add "03", "3학년"
'    oColumn.ValidValues.Add "04", "4학년"
'    oColumn.DisplayDesc = True
'
'    '매트릭스-회차
'    Set oColumn = oMat1.Columns("Count")
'    oColumn.ValidValues.Add "", ""
'    oColumn.ValidValues.Add "01", "1차"
'    oColumn.ValidValues.Add "02", "2차"
'    oColumn.DisplayDesc = True



    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
    
PH_PY309_CreateItems_Error:
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY309_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY309_EnableMenus()
    
    On Error GoTo PH_PY309_EnableMenus_Error
    
    Call oForm.EnableMenu("1283", False)         '// 삭제
    Call oForm.EnableMenu("1287", False)          '// 복제
'    Call oForm.EnableMenu("1286", True)         '// 닫기
    Call oForm.EnableMenu("1284", True)         '// 취소
    Call oForm.EnableMenu("1293", True)         '// 행삭제
    
    Exit Sub
    
PH_PY309_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY309_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY309_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY309_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY309_FormItemEnabled
        Call PH_PY309_AddMatrixRow
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY309_FormItemEnabled
        oForm.Items("DocEntry").Specific.Value = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY309_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY309_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY309_FormItemEnabled()
    On Error GoTo PH_PY309_FormItemEnabled_Error
    
    Dim oCombo As SAPbouiCOM.ComboBox
    
    Call oForm.Freeze(True)
    If (oForm.Mode = fm_ADD_MODE) Then
        oForm.Items("CLTCOD").Enabled = True
        oForm.Items("CntcCode").Enabled = True
        oForm.Items("LoanDate").Enabled = True
        oForm.Items("LoanAmt").Enabled = True
        oForm.Items("RpmtPrd").Enabled = True
        oForm.Items("IntRate").Enabled = True
        oForm.Items("DocEntry").Enabled = False
        oForm.Items("btnCal").Enabled = True
        
        '폼 DocEntry 세팅
        Call PH_PY309_FormClear
        
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        
        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", False)    '//문서추가
        
    ElseIf (oForm.Mode = fm_FIND_MODE) Then
        oForm.Items("CLTCOD").Enabled = True
        oForm.Items("CntcCode").Enabled = True
        oForm.Items("LoanDate").Enabled = True
        oForm.Items("LoanAmt").Enabled = True
        oForm.Items("RpmtPrd").Enabled = True
        oForm.Items("IntRate").Enabled = True
        oForm.Items("DocEntry").Enabled = True
        oForm.Items("btnCal").Enabled = True
        
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        
        Call oForm.EnableMenu("1281", False)    '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가
        
    ElseIf (oForm.Mode = fm_OK_MODE) Then
        oForm.Items("CLTCOD").Enabled = False
        oForm.Items("CntcCode").Enabled = False
        oForm.Items("LoanDate").Enabled = False
        oForm.Items("LoanAmt").Enabled = False
        oForm.Items("RpmtPrd").Enabled = False
        oForm.Items("IntRate").Enabled = False
        oForm.Items("DocEntry").Enabled = False
        oForm.Items("btnCal").Enabled = False
        
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD", False)
        
        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가

    End If
    Call oForm.Freeze(False)
    Exit Sub
    
PH_PY309_FormItemEnabled_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY309_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub Raise_FormItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
    On Error GoTo Raise_FormItemEvent_Error
    
    Dim sQry As String
    Dim i As Long
    Dim oCombo As SAPbouiCOM.ComboBox
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    Dim loopCount As Integer 'For Loop 용 (VALIDATE Event에서 사용)
    Dim GovID1 As String '주민등록번호 앞자리(VALIDATE Event에서 사용)
    Dim GovID2 As String '주민등록번호 뒷자리(VALIDATE Event에서 사용)
    Dim GovID As String '주민등록번호 전체(VALIDATE Event에서 사용)
    Dim Sex As String '성별(VALIDATE Event에서 사용)
    Dim SchCls As String '학교(VALIDATE Event에서 사용)
    Dim PayCnt As Integer '지급횟수(COMBO_SELECT Event에서 사용)
    Dim Tuition As Double '등록금계(VALIDATE Event에서 사용)
    Dim FeeTot As Double '입학금계(VALIDATE Event에서 사용)
    Dim TuiTot As Double '등록금계(VALIDATE Event에서 사용)
    Dim Total As Double '총계(VALIDATE Event에서 사용)
    
    Dim PreTuition As Double '등록금 입력 전 데이터
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Select Case pval.EventType
        Case et_ITEM_PRESSED: '//1
            
            If pval.BeforeAction = True Then
                If pval.ItemUID = "1" Then
                    If oForm.Mode = fm_ADD_MODE Then
                        If PH_PY309_DataValidCheck = False Then
                            BubbleEvent = False
                        End If
                        
                        '//해야할일 작업
                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
                        If PH_PY309_DataValidCheck = False Then
                            BubbleEvent = False
                        End If
                        '//해야할일 작업
                        
                    ElseIf oForm.Mode = fm_OK_MODE Then
                    End If
                    
                ElseIf pval.ItemUID = "btnCal" Then
                
                    If PH_PY309_CalDataCheck = False Then
                        BubbleEvent = False
                    Else
                        Call PH_PY309_MTX01
                    End If
                    
                End If
            ElseIf pval.BeforeAction = False Then
                If pval.ItemUID = "1" Then
                    If oForm.Mode = fm_ADD_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY309_FormItemEnabled
                            Call PH_PY309_AddMatrixRow
                        End If
                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY309_FormItemEnabled
                            Call PH_PY309_AddMatrixRow
                        End If
                    ElseIf oForm.Mode = fm_OK_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY309_FormItemEnabled
                        End If
                    End If
                End If
            End If
     
'----------------------------------------------------------
        Case et_KEY_DOWN: '//2

            If pval.BeforeAction = True Then
            
                If pval.ItemUID = "Mat01" Then
        
                    If pval.ColUID = "Name" And pval.CharPressed = "9" Then
'
'                        If oMat1.Columns.Item("Name").Cells(pval.Row).Specific.Value = "" Then
'                            Call Sbo_Application.ActivateMenuItem("7425")
'                            BubbleEvent = False
'                        End If
                    
                    End If
                
                ElseIf pval.ItemUID = "CntcCode" And pval.CharPressed = "9" Then
                
                    If oForm.Items("CntcCode").Specific.Value = "" Then
                        Call Sbo_Application.ActivateMenuItem("7425")
                        BubbleEvent = False
                    End If
            
                End If

            ElseIf pval.Before_Action = False Then

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
'----------------------------------------------------------
        Case et_CLICK: '//6
            If pval.BeforeAction = True Then
                Select Case pval.ItemUID
                Case "Mat01"
                    If pval.Row > 0 Then
'                        Call oMat1.SelectRow(pval.Row, True, False)
                    End If
                End Select
                
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

                    Select Case pval.ItemUID
                    
                        Case "Mat01"

                            If (PH_PY309_Validate("수정", pval.Row) = False) Then
                                Call oDS_PH_PY309B.setValue("U_" & pval.ColUID, pval.Row - 1, Trim(oDS_PH_PY309B.GetValue("U_" & pval.ColUID, pval.Row - 1)))
                                oMat1.LoadFromDataSource
                            End If
                    
                    End Select

                End If
       
            ElseIf pval.BeforeAction = False Then
            
                If pval.ItemChanged = True Then
                    
                    Select Case pval.ItemUID
                    
                        Case "CntcCode"
                        
                            Call oDS_PH_PY309A.setValue("U_CntcName", 0, MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & oForm.Items(pval.ItemUID).Specific.Value & "'"))
                        
                        Case "Mat01"
                                                
                            If pval.ColUID = "Cnt" Then
                            
                                oMat1.FlushToDataSource
                                
                                Call oDS_PH_PY309B.setValue("U_" & pval.ColUID, pval.Row - 1, oMat1.Columns(pval.ColUID).Cells(pval.Row).Specific.Value)
                                
                                oMat1.LoadFromDataSource
                                
                                If oMat1.RowCount = pval.Row And Trim(oDS_PH_PY309B.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
                                    Call PH_PY309_AddMatrixRow
                                End If
                                
                            End If
                            
                            Call oMat1.Columns(pval.ColUID).Cells(pval.Row).CLICK(ct_Regular)
                            Call oMat1.AutoResizeColumns
                            
                    End Select

                End If
                
            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_MATRIX_LOAD: '//11
            If pval.BeforeAction = True Then
            ElseIf pval.BeforeAction = False Then
                oMat1.LoadFromDataSource

                Call PH_PY309_FormItemEnabled
                Call PH_PY309_AddMatrixRow
                Call oMat1.AutoResizeColumns
                
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
                Set oDS_PH_PY309A = Nothing
                Set oDS_PH_PY309B = Nothing
                
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

                Call oMat1.AutoResizeColumns

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
'                    Call MDC_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY309A", "Code")
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
    Sbo_Application.SetStatusBarMessage "Raise_ItemEvent_Error: " & Err.Number & " - " & Err.Description & "EventType : " & pval.EventType, bmt_Short, True
End Sub


Public Sub Raise_FormMenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
    Dim i As Long
    On Error GoTo Raise_FormMenuEvent_Error
    
    Dim loopCount As Integer
    Dim FeeTot As Double
    Dim TuiTot As Double
    Dim Total As Double
    
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
            
                If (PH_PY309_Validate("행삭제") = False) Then
                    BubbleEvent = False
                    Call oForm.Freeze(False)
                    Exit Sub
                End If
            
            Case "1281":
            Case "1282":
            Case "1288", "1289", "1290", "1291":
                Call AuthorityCheck(oForm, "CLTCOD", "@PH_PY309A", "DocEntry")      '//접속자 권한에 따른 사업장 보기
        End Select
    ElseIf (pval.BeforeAction = False) Then
        Select Case pval.MenuUID
            Case "1283":
                oForm.Mode = fm_ADD_MODE
                Call PH_PY309_FormItemEnabled
                Call PH_PY309_AddMatrixRow
            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//문서찾기
                Call PH_PY309_FormItemEnabled
                Call PH_PY309_AddMatrixRow
                oForm.Items("DocEntry").CLICK ct_Regular
            Case "1282":    '//문서추가
                Call PH_PY309_FormItemEnabled
                Call PH_PY309_AddMatrixRow
            Case "1288", "1289", "1290", "1291":
                Call PH_PY309_FormItemEnabled
            Case "1293" '// 행삭제

                 If oMat1.RowCount <> oMat1.VisualRowCount Then
                    oMat1.FlushToDataSource
    
                    While (i <= oDS_PH_PY309B.Size - 1)
                        If oDS_PH_PY309B.GetValue("U_LineNum", i) = "" Then
                            oDS_PH_PY309B.RemoveRecord (i)
                            i = 0
                        Else
                            i = i + 1
                        End If
                    Wend
                    
                    For i = 0 To oDS_PH_PY309B.Size
                        Call oDS_PH_PY309B.setValue("U_LineNum", i, i + 1)
                    Next i
                    
                    oMat1.LoadFromDataSource
                End If
                
            '복제
            Case "1287"
            
                Call oForm.Freeze(True)
                Call oDS_PH_PY309A.setValue("DocEntry", 0, "")
                
                For i = 0 To oMat1.VisualRowCount - 1
                    Call oMat1.FlushToDataSource
                    Call oDS_PH_PY309B.setValue("DocEntry", i, "")
                    Call oDS_PH_PY309B.setValue("U_PayYN", i, "N")
                    Call oMat1.LoadFromDataSource
                Next i
                Call oForm.Freeze(False)
            
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

Sub PH_PY309_AddMatrixRow()
    Dim oRow    As Long
    
    On Error GoTo PH_PY309_AddMatrixRow_Error
    
    Call oForm.Freeze(True)
    
    '//[Mat1]
    oMat1.FlushToDataSource
    oRow = oMat1.VisualRowCount
    
    If oMat1.VisualRowCount > 0 Then
        If Trim(oDS_PH_PY309B.GetValue("U_LineNum", oRow - 1)) <> "" Then
            If oDS_PH_PY309B.Size <= oMat1.VisualRowCount Then
                oDS_PH_PY309B.InsertRecord (oRow)
            End If
            oDS_PH_PY309B.Offset = oRow
            oDS_PH_PY309B.setValue "U_LineNum", oRow, oRow + 1
            oDS_PH_PY309B.setValue "U_Cnt", oRow, ""
            oDS_PH_PY309B.setValue "U_PayDate", oRow, ""
            oDS_PH_PY309B.setValue "U_RpmtAmt", oRow, 0
            oDS_PH_PY309B.setValue "U_TotRpmt", oRow, 0
            oDS_PH_PY309B.setValue "U_RpmtYN", oRow, "N"
            oMat1.LoadFromDataSource
        Else
            oDS_PH_PY309B.Offset = oRow - 1
            oDS_PH_PY309B.setValue "U_LineNum", oRow - 1, oRow
            oDS_PH_PY309B.setValue "U_Cnt", oRow - 1, ""
            oDS_PH_PY309B.setValue "U_PayDate", oRow - 1, ""
            oDS_PH_PY309B.setValue "U_RpmtAmt", oRow - 1, 0
            oDS_PH_PY309B.setValue "U_TotRpmt", oRow - 1, 0
            oDS_PH_PY309B.setValue "U_RpmtYN", oRow - 1, "N"
            oMat1.LoadFromDataSource
        End If
    ElseIf oMat1.VisualRowCount = 0 Then
        oDS_PH_PY309B.Offset = oRow
        oDS_PH_PY309B.setValue "U_LineNum", oRow, oRow + 1
        oDS_PH_PY309B.setValue "U_Cnt", oRow, ""
        oDS_PH_PY309B.setValue "U_PayDate", oRow, ""
        oDS_PH_PY309B.setValue "U_RpmtAmt", oRow, 0
        oDS_PH_PY309B.setValue "U_TotRpmt", oRow, 0
        oDS_PH_PY309B.setValue "U_RpmtYN", oRow, "N"
        oMat1.LoadFromDataSource
    End If
    
    Call oForm.Freeze(False)
    Exit Sub
PH_PY309_AddMatrixRow_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY309_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY309_FormClear()
On Error GoTo PH_PY309_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY309'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.Value = 1
    Else
        oForm.Items("DocEntry").Specific.Value = DocEntry
    End If
    Exit Sub
PH_PY309_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY309_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY309_DataValidCheck() As Boolean
On Error GoTo PH_PY309_DataValidCheck_Error
    PH_PY309_DataValidCheck = False
    Dim i       As Long
    Dim sQry    As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    '사업장
    If Trim(oDS_PH_PY309A.GetValue("U_CLTCOD", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "사업장은 필수입니다.", bmt_Short, True
        oForm.Items("CLTCOD").CLICK ct_Regular
        PH_PY309_DataValidCheck = False
        Exit Function
    End If
    
    '사번
    If Trim(oDS_PH_PY309A.GetValue("U_CntcCode", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "사번은 필수입니다.", bmt_Short, True
        oForm.Items("CntcCode").CLICK ct_Regular
        PH_PY309_DataValidCheck = False
        Exit Function
    End If
    
    '대출일자
    If Trim(oDS_PH_PY309A.GetValue("U_LoanDate", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "대출일자는 필수입니다.", bmt_Short, True
        oForm.Items("LoanDate").CLICK ct_Regular
        PH_PY309_DataValidCheck = False
        Exit Function
    End If
    
    '대출금액
    If Trim(oDS_PH_PY309A.GetValue("U_LoanAmt", 0)) = 0 Then
        Sbo_Application.SetStatusBarMessage "대출금액은 필수입니다.", bmt_Short, True
        oForm.Items("LoanAmt").CLICK ct_Regular
        PH_PY309_DataValidCheck = False
        Exit Function
    End If
    
    '상환기간
    If Trim(oDS_PH_PY309A.GetValue("U_RpmtPrd", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "상환기간은 필수입니다.", bmt_Short, True
        oForm.Items("RpmtPrd").CLICK ct_Regular
        PH_PY309_DataValidCheck = False
        Exit Function
    End If
    
    '라인
    If oMat1.VisualRowCount > 1 Then
'        For i = 1 To oMat1.VisualRowCount - 1
'
'            '학교
'            If oMat1.Columns("SchCls").Cells(i).Specific.Value = "" Then
'                Sbo_Application.SetStatusBarMessage "학교는 필수입니다.", bmt_Short, True
'                oMat1.Columns("SchCls").Cells(i).CLICK ct_Regular
'                PH_PY309_DataValidCheck = False
'                Exit Function
'            End If
'
'            '학교명
'            If oMat1.Columns("SchName").Cells(i).Specific.Value = "" Then
'                Sbo_Application.SetStatusBarMessage "학교명은 필수입니다.", bmt_Short, True
'                oMat1.Columns("SchName").Cells(i).CLICK ct_Regular
'                PH_PY309_DataValidCheck = False
'                Exit Function
'            End If
'
'            '학년
'            If oMat1.Columns("Grade").Cells(i).Specific.Value = "" Then
'                Sbo_Application.SetStatusBarMessage "학년은 필수입니다.", bmt_Short, True
'                oMat1.Columns("Grade").Cells(i).CLICK ct_Regular
'                PH_PY309_DataValidCheck = False
'                Exit Function
'            End If
'
'            '회차
'            If oMat1.Columns("Count").Cells(i).Specific.Value = "" Then
'                Sbo_Application.SetStatusBarMessage "회차는 필수입니다.", bmt_Short, True
'                oMat1.Columns("Count").Cells(i).CLICK ct_Regular
'                PH_PY309_DataValidCheck = False
'                Exit Function
'            End If
'
'        Next
    Else
        Sbo_Application.SetStatusBarMessage "라인 데이터가 없습니다.", bmt_Short, True
        PH_PY309_DataValidCheck = False
        Exit Function
    End If
    
        oMat1.FlushToDataSource
    '// Matrix 마지막 행 삭제(DB 저장시)
    If oDS_PH_PY309B.Size > 1 Then oDS_PH_PY309B.RemoveRecord (oDS_PH_PY309B.Size - 1)
    
    oMat1.LoadFromDataSource
    
    PH_PY309_DataValidCheck = True
    
    Exit Function
    
    Set oRecordSet = Nothing
    
    
PH_PY309_DataValidCheck_Error:
    Set oRecordSet = Nothing
    PH_PY309_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY309_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY309_MTX01()
On Error GoTo PH_PY309_MTX01_Error
    '//메트릭스에 데이터 로드
    
    Dim i As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Dim Param01 As String
    Dim Param02 As String
    Dim Param03 As String
    Dim Param04 As String
    
    Param01 = oForm.Items("LoanAmt").Specific.Value
    Param02 = oForm.Items("LoanDate").Specific.Value
    Param03 = oForm.Items("RpmtPrd").Specific.Value
'    Param04 = oForm.Items("Param01").Specific.Value
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet.RecordCount, False)

    Call oForm.Freeze(True)
    
    sQry = "EXEC PH_PY309_01 '" & Param01 & "','" & Param02 & "','" & Param03 & "'"
    Call oRecordSet.DoQuery(sQry)
    
    oMat1.Clear
    oMat1.FlushToDataSource
    oMat1.LoadFromDataSource
    
    If (oRecordSet.RecordCount = 0) Then
        Call oMat1.Clear
        GoTo PH_PY309_MTX01_Exit
    End If

    For i = 0 To oRecordSet.RecordCount
        If i <> 0 Then
            Call oDS_PH_PY309B.InsertRecord(i)
        End If
        
        If i = oRecordSet.RecordCount Then '마지막 빈행 추가
        
            oDS_PH_PY309B.Offset = i
            Call oDS_PH_PY309B.setValue("U_LineNum", i, i + 1) '라인번호
            Call oDS_PH_PY309B.setValue("U_Cnt", i, "")
            Call oDS_PH_PY309B.setValue("U_PayDate", i, "")
            Call oDS_PH_PY309B.setValue("U_RpmtAmt", i, 0)
            Call oDS_PH_PY309B.setValue("U_TotRpmt", i, 0)
        
        Else
        
            oDS_PH_PY309B.Offset = i
            Call oDS_PH_PY309B.setValue("U_LineNum", i, i + 1) '라인번호
            Call oDS_PH_PY309B.setValue("U_Cnt", i, oRecordSet.Fields("Cnt").Value)
            Call oDS_PH_PY309B.setValue("U_PayDate", i, oRecordSet.Fields("PayDate").Value)
            Call oDS_PH_PY309B.setValue("U_RpmtAmt", i, oRecordSet.Fields("RpmtAmt").Value)
            Call oDS_PH_PY309B.setValue("U_TotRpmt", i, oRecordSet.Fields("TotRpmt").Value)
            
            oRecordSet.MoveNext
        
        End If
        
        ProgressBar01.Value = ProgressBar01.Value + 1
        ProgressBar01.Text = ProgressBar01.Value & "/" & oRecordSet.RecordCount & "건 조회중...!"
        
    Next
    
    oMat1.LoadFromDataSource
    oMat1.AutoResizeColumns
    oForm.Update
    
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Sub
PH_PY309_MTX01_Exit:
    Set oRecordSet = Nothing
    Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
    Call oForm.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Exit Sub
PH_PY309_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY309_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY309_Validate(ByVal ValidateType As String, Optional ByVal prmRow As Integer) As Boolean
On Error GoTo PH_PY309_Validate_Error
    PH_PY309_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY309A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.Value & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", bmt_Short, True
        PH_PY309_Validate = False
        GoTo PH_PY309_Validate_Exit
    End If
'
    If ValidateType = "수정" Then
    
        If oDS_PH_PY309B.GetValue("U_RpmtYN", prmRow - 1) = "Y" Then
        
            MDC_Com.MDC_GF_Message "상환이 완료된 행입니다. 수정할 수 없습니다.", "W"
            PH_PY309_Validate = False
            GoTo PH_PY309_Validate_Exit
        
        End If

    ElseIf ValidateType = "행삭제" Then

        If oDS_PH_PY309B.GetValue("U_RpmtYN", oLastColRow - 1) = "Y" Then
        
            MDC_Com.MDC_GF_Message "상환이 완료된 행입니다. 삭제할 수 없습니다.", "W"
            PH_PY309_Validate = False
            GoTo PH_PY309_Validate_Exit
        
        End If

    ElseIf ValidateType = "취소" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY309_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY309_Validate_Error:
    PH_PY309_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY309_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Function PH_PY309_CalDataCheck() As Boolean
'******************************************************************************
'Function ID : PH_PY309_CalDataCheck()
'해당모듈 : PH_PY309
'기능 : 상환스케줄 계산 시 필수 데이터 체크
'인수 : 없음
'반환값 : True : 필수 데이터 전체 다 입력했으면, False : 필수 데이터 중 하나라도 입력이 되지 않으면
'특이사항 : 없음
'******************************************************************************
On Error GoTo PH_PY309_CalDataCheck_Error

    PH_PY309_CalDataCheck = False
    
    Dim i As Long
    Dim sQry As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    '사업장
    If Trim(oDS_PH_PY309A.GetValue("U_CLTCOD", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "사업장은 필수입니다.", bmt_Short, True
        oForm.Items("CLTCOD").CLICK ct_Regular
        PH_PY309_CalDataCheck = False
        Exit Function
    End If
    
    '사번
    If Trim(oDS_PH_PY309A.GetValue("U_CntcCode", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "사번은 필수입니다.", bmt_Short, True
        oForm.Items("CntcCode").CLICK ct_Regular
        PH_PY309_CalDataCheck = False
        Exit Function
    End If
    
    '대출일자
    If Trim(oDS_PH_PY309A.GetValue("U_LoanDate", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "대출일자는 필수입니다.", bmt_Short, True
        oForm.Items("LoanDate").CLICK ct_Regular
        PH_PY309_CalDataCheck = False
        Exit Function
    End If
    
    '대출금액
    If Trim(oDS_PH_PY309A.GetValue("U_LoanAmt", 0)) = 0 Then
        Sbo_Application.SetStatusBarMessage "대출금액은 필수입니다.", bmt_Short, True
        oForm.Items("LoanAmt").CLICK ct_Regular
        PH_PY309_CalDataCheck = False
        Exit Function
    End If
    
    '이자율
    If Trim(oDS_PH_PY309A.GetValue("U_IntRate", 0)) = "0.0" Then
        Sbo_Application.SetStatusBarMessage "이자율은 필수입니다.", bmt_Short, True
        oForm.Items("IntRate").CLICK ct_Regular
        PH_PY309_CalDataCheck = False
        Exit Function
    End If
    
    '상환기간
    If Trim(oDS_PH_PY309A.GetValue("U_RpmtPrd", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "상환기간은 필수입니다.", bmt_Short, True
        oForm.Items("RpmtPrd").CLICK ct_Regular
        PH_PY309_CalDataCheck = False
        Exit Function
    End If
    
    PH_PY309_CalDataCheck = True
    
    Exit Function
    
    Set oRecordSet = Nothing
        
PH_PY309_CalDataCheck_Error:
    Set oRecordSet = Nothing
    PH_PY309_CalDataCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY309_CalDataCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function
