VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY006"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY006.cls
'//  Module         : 인사관리 > 인사마스터관리
'//  Desc           : 승호등록작업
'//********************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form

Public oMat1 As SAPbouiCOM.Matrix

Private oDS_PH_PY006A As SAPbouiCOM.DBDataSource
Private oDS_PH_PY006B As SAPbouiCOM.DBDataSource

Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long

Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY006.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY006_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY006"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
    oForm.DataBrowser.BrowseBy = "DocEntry"
        
    oForm.Freeze True
    Call PH_PY006_CreateItems
    Call PH_PY006_EnableMenus
    Call PH_PY006_SetDocument(oFromDocEntry01)
'    Call PH_PY006_FormResize
    
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

Private Function PH_PY006_CreateItems() As Boolean

    Dim sQry        As String
    Dim i           As Long
    
    Dim oEdit       As SAPbouiCOM.EditText
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns    As SAPbouiCOM.Columns
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY006_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Set oDS_PH_PY006A = oForm.DataSources.DBDataSources("@PH_PY006A")
    Set oDS_PH_PY006B = oForm.DataSources.DBDataSources("@PH_PY006B")

    
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
    
    '직원구분
    Set oCombo = oForm.Items("JIGTYP").Specific
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P126' AND U_UseYN= 'Y' "
    Call SetReDataCombo(oForm, sQry, oCombo)
    oForm.Items("JIGTYP").DisplayDesc = True
    
    '승호구분
    Set oCombo = oForm.Items("Gubun").Specific
    oCombo.ValidValues.Add "10", "금액승호(전문직)"
    oCombo.ValidValues.Add "20", "호봉승호(사무기술직)"
    oForm.Items("JIGTYP").DisplayDesc = True
    
    
    '매트릭스-부서
    Set oColumn = oMat1.Columns("TeamCode")
    sQry = "            SELECT      T1.U_Code,"
    sQry = sQry & "                 T1.U_CodeNm"
    sQry = sQry & "  FROM       [@PS_HR200H] AS T0"
    sQry = sQry & "                 INNER JOIN"
    sQry = sQry & "                 [@PS_HR200L] AS T1"
    sQry = sQry & "                     ON T0.Code = T1.Code"
    sQry = sQry & "  WHERE      T0.Code = '1'"
    sQry = sQry & "                 AND T1.U_UseYN = 'Y'"
    sQry = sQry & "  ORDER BY  T1.U_Seq"
    
    If oRecordSet.RecordCount > 0 Then
        Do Until oRecordSet.EOF
            oColumn.ValidValues.Add Trim$(oRecordSet.Fields(0).VALUE), Trim$(oRecordSet.Fields(1).VALUE)
            oRecordSet.MoveNext
        Loop
    End If
    
'    Call MDC_SetMod.GP_MatrixSetMatComboList(oColumn, sQry, False, False)
    oColumn.DisplayDesc = True
    
    '매트릭스-담당
    Set oColumn = oMat1.Columns("RspCode")
    sQry = "            SELECT      T1.U_Code,"
    sQry = sQry & "                 T1.U_CodeNm"
    sQry = sQry & "  FROM       [@PS_HR200H] AS T0"
    sQry = sQry & "                 INNER JOIN"
    sQry = sQry & "                 [@PS_HR200L] AS T1"
    sQry = sQry & "                     ON T0.Code = T1.Code"
    sQry = sQry & "  WHERE      T0.Code = '1'"
    sQry = sQry & "                 AND T1.U_UseYN = 'Y'"
    sQry = sQry & "  ORDER BY  T1.U_Seq"
    
    If oRecordSet.RecordCount > 0 Then
        Do Until oRecordSet.EOF
            oColumn.ValidValues.Add Trim$(oRecordSet.Fields(0).VALUE), Trim$(oRecordSet.Fields(1).VALUE)
            oRecordSet.MoveNext
        Loop
    End If
    
'    Call MDC_SetMod.GP_MatrixSetMatComboList(oColumn, sQry, False, False)
    oColumn.DisplayDesc = True
    
    '매트릭스-반
    Set oColumn = oMat1.Columns("ClsCode")
    sQry = "            SELECT      T1.U_Code,"
    sQry = sQry & "                 T1.U_CodeNm"
    sQry = sQry & "  FROM       [@PS_HR200H] AS T0"
    sQry = sQry & "                 INNER JOIN"
    sQry = sQry & "                 [@PS_HR200L] AS T1"
    sQry = sQry & "                     ON T0.Code = T1.Code"
    sQry = sQry & "  WHERE      T0.Code = '9'"
    sQry = sQry & "                 AND T1.U_UseYN = 'Y'"
    sQry = sQry & "  ORDER BY  T1.U_Seq"
    
    If oRecordSet.RecordCount > 0 Then
        Do Until oRecordSet.EOF
            oColumn.ValidValues.Add Trim$(oRecordSet.Fields(0).VALUE), Trim$(oRecordSet.Fields(1).VALUE)
            oRecordSet.MoveNext
        Loop
    End If
    
'    Call MDC_SetMod.GP_MatrixSetMatComboList(oColumn, sQry, False, False)
    oColumn.DisplayDesc = True
    
    '매트릭스-직급
    Set oColumn = oMat1.Columns("JIGCOD")
    sQry = "            SELECT      T1.U_Code,"
    sQry = sQry & "                 T1.U_CodeNm"
    sQry = sQry & "  FROM       [@PS_HR200H] AS T0"
    sQry = sQry & "                 INNER JOIN"
    sQry = sQry & "                 [@PS_HR200L] AS T1"
    sQry = sQry & "                     ON T0.Code = T1.Code"
    sQry = sQry & "  WHERE      T0.Code = 'P129'"
    sQry = sQry & "                 AND T1.U_UseYN = 'Y'"
    sQry = sQry & "  ORDER BY  T1.U_Seq"
    
    If oRecordSet.RecordCount > 0 Then
        Do Until oRecordSet.EOF
            oColumn.ValidValues.Add Trim$(oRecordSet.Fields(0).VALUE), Trim$(oRecordSet.Fields(1).VALUE)
            oRecordSet.MoveNext
        Loop
    End If
    
'    Call MDC_SetMod.GP_MatrixSetMatComboList(oColumn, sQry, False, False)
    oColumn.DisplayDesc = True
    
    '승호대상자 여부
    Set oColumn = oMat1.Columns("APPLYYN")
    oColumn.ValidValues.Add "Y", "승호대상자"
    oColumn.ValidValues.Add "N", "승호제외자"
    oColumn.DisplayDesc = True

    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
    
PH_PY006_CreateItems_Error:
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY006_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY006_EnableMenus()
    
    On Error GoTo PH_PY006_EnableMenus_Error
    
    Call oForm.EnableMenu("1283", False)         '// 삭제
    Call oForm.EnableMenu("1287", True)          '// 복제
'    Call oForm.EnableMenu("1286", True)         '// 닫기
    Call oForm.EnableMenu("1284", True)         '// 취소
    Call oForm.EnableMenu("1293", True)         '// 행삭제
    
    Exit Sub
    
PH_PY006_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY006_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY006_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY006_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY006_FormItemEnabled
        Call PH_PY006_AddMatrixRow
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY006_FormItemEnabled
        oForm.Items("DocEntry").Specific.VALUE = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY006_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY006_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY006_FormItemEnabled()
    On Error GoTo PH_PY006_FormItemEnabled_Error
    
    Dim oCombo As SAPbouiCOM.ComboBox
    
    Call oForm.Freeze(True)
    If (oForm.Mode = fm_ADD_MODE) Then
        oForm.Items("CLTCOD").Enabled = True
        oForm.Items("YM").Enabled = True
        oForm.Items("DocDate").Enabled = True
        oForm.Items("DocDate").Enabled = True
        oForm.Items("UPAMT").Enabled = True
        oForm.Items("appNum").Enabled = True
        oForm.Items("DocEntry").Enabled = False
        
        '폼 DocEntry 세팅
        Call PH_PY006_FormClear
        
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        
        '년도 세팅
        Call oDS_PH_PY006A.setValue("U_YM", 0, Format(Date, "YYYYMM"))

        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", False)    '//문서추가
        
    ElseIf (oForm.Mode = fm_FIND_MODE) Then
        oForm.Items("CLTCOD").Enabled = True
        oForm.Items("YM").Enabled = True
        oForm.Items("DocDate").Enabled = True
        oForm.Items("UPAMT").Enabled = True
        oForm.Items("JIGTYP").Enabled = True
        oForm.Items("appNum").Enabled = True
        oForm.Items("DocEntry").Enabled = True
        
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        
        Call oForm.EnableMenu("1281", False)    '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가
        
    ElseIf (oForm.Mode = fm_OK_MODE) Then
        oForm.Items("CLTCOD").Enabled = False
        oForm.Items("YM").Enabled = False
        oForm.Items("DocDate").Enabled = False
        oForm.Items("UPAMT").Enabled = False
        oForm.Items("JIGTYP").Enabled = False
        oForm.Items("appNum").Enabled = False
        oForm.Items("DocEntry").Enabled = False
        
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD", False)
        
'        Call oForm.EnableMenu("1281", True)     '//문서찾기
'        Call oForm.EnableMenu("1282", True)     '//문서추가

    End If
    Call oForm.Freeze(False)
    Exit Sub
    
PH_PY006_FormItemEnabled_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY006_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub Raise_FormItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
    On Error GoTo Raise_FormItemEvent_Error
    
    Dim sQry As String
    Dim i As Long
    Dim oCombo As SAPbouiCOM.ComboBox
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    
    Dim PreTuition As Double '등록금 입력 전 데이터
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Select Case pval.EventType
        Case et_ITEM_PRESSED: '//1
            
            If pval.BeforeAction = True Then
                If pval.ItemUID = "1" Then
                    If oForm.Mode = fm_ADD_MODE Then
                        If PH_PY006_DataValidCheck = False Then
                            BubbleEvent = False
                        End If
                        
                        '//해야할일 작업
                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
                        If PH_PY006_DataValidCheck = False Then
                            BubbleEvent = False
                        End If
                        '//해야할일 작업
                        
                    ElseIf oForm.Mode = fm_OK_MODE Then
                    End If
                    
                ElseIf pval.ItemUID = "Btn01" Then
                    '대상자조회
                    Call PH_PY006_MTX01
                    
                ElseIf pval.ItemUID = "Btn02" Then
                    '승호처리 및 발령사항 추가
                    If oForm.Mode = fm_OK_MODE Then
                        If oForm.Items("Canceled").Specific.VALUE = "N" Then
                            If Trim(oForm.Items("appNum").Specific.VALUE) = "" Then
                                Sbo_Application.SetStatusBarMessage "발령번호가 없습니다..", bmt_Short, True
                            Else
                                Call PH_PY006_MTX02
                            End If
                        Else
                            Sbo_Application.SetStatusBarMessage "취소된 문서는 처리할 수 없습니다..", bmt_Short, True
                        End If
                    Else
                        Sbo_Application.SetStatusBarMessage "추가(수정)후 처리바랍니다.", bmt_Short, True
                    End If
                ElseIf pval.ItemUID = "Btn03" Then
                    '승호취소 처리
                    If oForm.Mode = fm_OK_MODE Then
                        If oForm.Items("Canceled").Specific.VALUE = "Y" Then
                            Call PH_PY006_MTX03
                        Else
                            Sbo_Application.SetStatusBarMessage "취소처리된 문서만 취소처리할 수 있습니다..", bmt_Short, True
                        End If
                    Else
                        Sbo_Application.SetStatusBarMessage "추가(갱신)후 처리바랍니다.", bmt_Short, True
                    End If
                End If
            ElseIf pval.BeforeAction = False Then
                If pval.ItemUID = "1" Then
                    If oForm.Mode = fm_ADD_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY006_FormItemEnabled
                            Call PH_PY006_AddMatrixRow
                        End If
                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY006_FormItemEnabled
                            Call PH_PY006_AddMatrixRow
                        End If
                    ElseIf oForm.Mode = fm_OK_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY006_FormItemEnabled
                        End If
                    End If
                End If
            End If
     
'----------------------------------------------------------
        Case et_KEY_DOWN: '//2

            If pval.BeforeAction = True Then
            
                If pval.ItemUID = "Mat01" Then
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
'                    Call PH_PY006_AddMatrixRow
                    
                    If pval.ItemUID = "JIGTYP" Or pval.ItemUID = "CLTCOD" Then
                        oMat1.Clear
                    ElseIf pval.ItemUID = "Mat01" Then
                        If pval.ColUID = "" Then
                        
                            Call oMat1.FlushToDataSource
                            
                            
                            Call oMat1.LoadFromDataSource
                        
                        End If
                        
                        Call oMat1.AutoResizeColumns
                    End If
                    
                End If
            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_CLICK: '//6
            If pval.BeforeAction = True Then
                Select Case pval.ItemUID
                Case "Mat01"
                    If pval.Row > 0 Then
                        Call oMat1.SelectRow(pval.Row, True, False)
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

                End If
       
            ElseIf pval.BeforeAction = False Then
            
                If pval.ItemChanged = True Then
                    
                    Select Case pval.ItemUID
                    
                        Case "CntcCode"
                            Call oDS_PH_PY006A.setValue("U_CntcName", 0, MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & oForm.Items(pval.ItemUID).Specific.VALUE & "'"))
                        
                        Case "Mat01"
                        
                            If pval.ColUID = "" Then
                            
                                oMat1.FlushToDataSource
                                
                                oMat1.LoadFromDataSource
                                
                                Call PH_PY006_AddMatrixRow
                            
                            
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

                Call PH_PY006_FormItemEnabled
                Call PH_PY006_AddMatrixRow
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
                Set oDS_PH_PY006A = Nothing
                Set oDS_PH_PY006B = Nothing
                
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
'                    Call MDC_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY006A", "Code")
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
    Sbo_Application.SetStatusBarMessage "Raise_ItemEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
            Case "1281":
            Case "1282":
            Case "1288", "1289", "1290", "1291":
                Call AuthorityCheck(oForm, "CLTCOD", "@PH_PY006A", "DocEntry")      '//접속자 권한에 따른 사업장 보기
        End Select
    ElseIf (pval.BeforeAction = False) Then
        Select Case pval.MenuUID
            Case "1283":
                oForm.Mode = fm_ADD_MODE
                Call PH_PY006_FormItemEnabled
                Call PH_PY006_AddMatrixRow
            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//문서찾기
                Call PH_PY006_FormItemEnabled
                Call PH_PY006_AddMatrixRow
                oForm.Items("DocEntry").CLICK ct_Regular
            Case "1282":    '//문서추가
                Call PH_PY006_FormItemEnabled
                Call PH_PY006_AddMatrixRow
            Case "1288", "1289", "1290", "1291":
                Call PH_PY006_FormItemEnabled
            Case "1293" '// 행삭제

                
                
            '복제
            Case "1287"
            
                
            
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

Sub PH_PY006_AddMatrixRow()
    Dim oRow    As Long
    
    On Error GoTo PH_PY006_AddMatrixRow_Error
    
    Call oForm.Freeze(True)
    
    '//[Mat1]
    oMat1.FlushToDataSource
    oRow = oMat1.VisualRowCount
    
'    If oMat1.VisualRowCount > 0 Then
'        If Trim(oDS_PH_PY006B.GetValue("U_Name", oRow - 1)) <> "" Then
'            If oDS_PH_PY006B.Size <= oMat1.VisualRowCount Then
'                oDS_PH_PY006B.InsertRecord (oRow)
'            End If
'            oDS_PH_PY006B.Offset = oRow
'            oDS_PH_PY006B.setValue "U_LineNum", oRow, oRow + 1
'            oDS_PH_PY006B.setValue "U_Name", oRow, ""
'            oDS_PH_PY006B.setValue "U_GovID", oRow, ""
'            oDS_PH_PY006B.setValue "U_Sex", oRow, ""
'            oDS_PH_PY006B.setValue "U_SchCls", oRow, ""
'            oDS_PH_PY006B.setValue "U_SchName", oRow, ""
'            oDS_PH_PY006B.setValue "U_Grade", oRow, ""
'            oDS_PH_PY006B.setValue "U_EntFee", oRow, 0
'            oDS_PH_PY006B.setValue "U_Tuition", oRow, 0
'            oDS_PH_PY006B.setValue "U_Count", oRow, ""
'            oDS_PH_PY006B.setValue "U_PayCnt", oRow, ""
'            oDS_PH_PY006B.setValue "U_PayYN", oRow, ""
'            oMat1.LoadFromDataSource
'        Else
'            oDS_PH_PY006B.Offset = oRow - 1
'            oDS_PH_PY006B.setValue "U_LineNum", oRow - 1, oRow
'            oDS_PH_PY006B.setValue "U_Name", oRow - 1, ""
'            oDS_PH_PY006B.setValue "U_GovID", oRow - 1, ""
'            oDS_PH_PY006B.setValue "U_Sex", oRow - 1, ""
'            oDS_PH_PY006B.setValue "U_SchCls", oRow - 1, ""
'            oDS_PH_PY006B.setValue "U_SchName", oRow - 1, ""
'            oDS_PH_PY006B.setValue "U_Grade", oRow - 1, ""
'            oDS_PH_PY006B.setValue "U_EntFee", oRow - 1, 0
'            oDS_PH_PY006B.setValue "U_Tuition", oRow - 1, 0
'            oDS_PH_PY006B.setValue "U_Count", oRow - 1, ""
'            oDS_PH_PY006B.setValue "U_PayCnt", oRow, ""
'            oDS_PH_PY006B.setValue "U_PayYN", oRow - 1, ""
'            oMat1.LoadFromDataSource
'        End If
'    ElseIf oMat1.VisualRowCount = 0 Then
'        oDS_PH_PY006B.Offset = oRow
'        oDS_PH_PY006B.setValue "U_LineNum", oRow, oRow + 1
'        oDS_PH_PY006B.setValue "U_Name", oRow, ""
'        oDS_PH_PY006B.setValue "U_GovID", oRow, ""
'        oDS_PH_PY006B.setValue "U_Sex", oRow, ""
'        oDS_PH_PY006B.setValue "U_SchCls", oRow, ""
'        oDS_PH_PY006B.setValue "U_SchName", oRow, ""
'        oDS_PH_PY006B.setValue "U_Grade", oRow, ""
'        oDS_PH_PY006B.setValue "U_EntFee", oRow, 0
'        oDS_PH_PY006B.setValue "U_Tuition", oRow, 0
'        oDS_PH_PY006B.setValue "U_Count", oRow, ""
'        oDS_PH_PY006B.setValue "U_PayCnt", oRow, ""
'        oDS_PH_PY006B.setValue "U_PayYN", oRow, ""
'        oMat1.LoadFromDataSource
'    End If
    
    Call oForm.Freeze(False)
    Exit Sub
PH_PY006_AddMatrixRow_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY006_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY006_FormClear()
On Error GoTo PH_PY006_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY006'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PH_PY006_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY006_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY006_DataValidCheck() As Boolean
On Error GoTo PH_PY006_DataValidCheck_Error
    PH_PY006_DataValidCheck = False
    Dim i       As Long
    Dim sQry    As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    '사업장
    If Trim(oDS_PH_PY006A.GetValue("U_CLTCOD", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "사업장은 필수입니다.", bmt_Short, True
        oForm.Items("CLTCOD").CLICK ct_Regular
        PH_PY006_DataValidCheck = False
        Exit Function
    End If
    
    '년월
    If Trim(oDS_PH_PY006A.GetValue("U_YM", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "년월은 필수입니다.", bmt_Short, True
        oForm.Items("YM").CLICK ct_Regular
        PH_PY006_DataValidCheck = False
        Exit Function
    End If
    
    '승호일
    If Trim(oDS_PH_PY006A.GetValue("U_DocDate", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "승호일은 필수입니다.", bmt_Short, True
        oForm.Items("DocDate").CLICK ct_Regular
        PH_PY006_DataValidCheck = False
        Exit Function
    End If
    
    '승호금액
    If Trim(oDS_PH_PY006A.GetValue("U_UPAMT", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "승호금액은 필수입니다.", bmt_Short, True
        oForm.Items("UPAMT").CLICK ct_Regular
        PH_PY006_DataValidCheck = False
        Exit Function
    End If
    
    '직원구분
    If Trim(oDS_PH_PY006A.GetValue("U_JIGTYP", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "직원구분은 필수입니다.", bmt_Short, True
        oForm.Items("JIGTYP").CLICK ct_Regular
        PH_PY006_DataValidCheck = False
        Exit Function
    End If
    
     '발령번호
    If Trim(oDS_PH_PY006A.GetValue("U_appNum", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "발령번호는 필수입니다.", bmt_Short, True
        oForm.Items("appNum").CLICK ct_Regular
        PH_PY006_DataValidCheck = False
        Exit Function
    End If
    
    '라인
    If oMat1.VisualRowCount > 1 Then
'        For i = 1 To oMat1.VisualRowCount - 1
'
'        Next
    Else
        Sbo_Application.SetStatusBarMessage "라인 데이터가 없습니다.", bmt_Short, True
        PH_PY006_DataValidCheck = False
        Exit Function
    End If
    
        oMat1.FlushToDataSource
    '// Matrix 마지막 행 삭제(DB 저장시)
'    If oDS_PH_PY006B.Size > 1 Then oDS_PH_PY006B.RemoveRecord (oDS_PH_PY006B.Size - 1)
    
    oMat1.LoadFromDataSource
    
    PH_PY006_DataValidCheck = True
    
    Exit Function
    
    Set oRecordSet = Nothing
    
    
PH_PY006_DataValidCheck_Error:
    Set oRecordSet = Nothing
    PH_PY006_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY006_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY006_MTX01()

    '//메트릭스에 데이터 로드
    
    Dim i       As Long
    Dim sQry    As String
    
    Dim YM      As String
    Dim DocDate As String
    
    Dim Param01 As String
    Dim Param02 As String
    Dim Param03 As String
    Dim Param04 As Double
    Dim Param05 As Double
    Dim Param06 As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY006_MTX01_Error

    Call oForm.Freeze(True)
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    

    Param01 = Trim(oForm.Items("CLTCOD").Specific.VALUE)
    Param02 = oForm.Items("JIGTYP").Specific.VALUE
    Param03 = oForm.Items("Gubun").Specific.VALUE
    Param04 = oForm.Items("UPHOBONG").Specific.VALUE
    Param05 = oForm.Items("UPAMT").Specific.VALUE
    Param06 = oForm.Items("DocDate").Specific.VALUE
    
    YM = oForm.Items("YM").Specific.VALUE
    DocDate = Trim(oForm.Items("DocDate").Specific.VALUE)
    
    
    If Trim(Param03) = "" Then
        Call MDC_Com.MDC_GF_Message("승호기준은 필수입니다. 확인바랍니다.", "E")
        oForm.Items("Gubun").CLICK ct_Regular
        GoTo PH_PY006_MTX01_Exit
    End If
    
    If Trim(Param06) = "" Then
        Call MDC_Com.MDC_GF_Message("승호일자는 필수입니다. 확인바랍니다.", "E")
        oForm.Items("DocDate").CLICK ct_Regular
        GoTo PH_PY006_MTX01_Exit
    End If
    
    
    sQry = "Select Count(*) From [@PH_PY006A] Where U_CLTCOD = '" & Param01 & "' and U_YM = '" & YM & "' and U_JIGTYP = '" & Param02 & "'"
    sQry = sQry + " and U_DocDate = '" & DocDate & "'"
    sQry = sQry + " and Canceled = 'N' "
    Call oRecordSet.DoQuery(sQry)
    If (oRecordSet.Fields(0).VALUE > 0) Then
        Call MDC_Com.MDC_GF_Message("승호작업을 이미 등록했습니다. 확인바랍니다.", "E")
        GoTo PH_PY006_MTX01_Exit
    End If
    
    sQry = "EXEC [PH_PY006_01] '" & Param01 & "', '" & Param02 & "', '" & Param03 & "', " & Param04 & ", " & Param05 & ", '" & Param06 & "'"
    Call oRecordSet.DoQuery(sQry)
    
    oMat1.Clear
    oMat1.FlushToDataSource
    oMat1.LoadFromDataSource
    
    If (oRecordSet.RecordCount = 0) Then
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
        GoTo PH_PY006_MTX01_Exit
    End If
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet.RecordCount, False)
    
    For i = 0 To oRecordSet.RecordCount - 1
        If i <> 0 Then
            oDS_PH_PY006B.InsertRecord (i)
        End If
        oDS_PH_PY006B.Offset = i
        oDS_PH_PY006B.setValue "U_LineNum", i, i + 1
        Call oDS_PH_PY006B.setValue("U_TeamCode", i, oRecordSet.Fields(0).VALUE)
        Call oDS_PH_PY006B.setValue("U_RspCode", i, oRecordSet.Fields(1).VALUE)
        Call oDS_PH_PY006B.setValue("U_ClsCode", i, oRecordSet.Fields(2).VALUE)
        Call oDS_PH_PY006B.setValue("U_MSTCOD", i, oRecordSet.Fields(3).VALUE)
        Call oDS_PH_PY006B.setValue("U_FULLNAME", i, oRecordSet.Fields(4).VALUE)
        Call oDS_PH_PY006B.setValue("U_JIGCOD", i, oRecordSet.Fields(5).VALUE)
        Call oDS_PH_PY006B.setValue("U_GrpDat", i, oRecordSet.Fields(6).VALUE)
        Call oDS_PH_PY006B.setValue("U_birthDat", i, oRecordSet.Fields(7).VALUE)
        Call oDS_PH_PY006B.setValue("U_HOBYMM", i, oRecordSet.Fields(8).VALUE)
        Call oDS_PH_PY006B.setValue("U_HOBONG", i, oRecordSet.Fields(9).VALUE)
        Call oDS_PH_PY006B.setValue("U_HOBNAM", i, oRecordSet.Fields(10).VALUE)
        Call oDS_PH_PY006B.setValue("U_STDAMT", i, oRecordSet.Fields(11).VALUE)
        Call oDS_PH_PY006B.setValue("U_BNSAMT", i, oRecordSet.Fields(12).VALUE)
        Call oDS_PH_PY006B.setValue("U_UPHOBONG", i, oRecordSet.Fields(13).VALUE)
        Call oDS_PH_PY006B.setValue("U_UPHOBNAM", i, oRecordSet.Fields(14).VALUE)
        Call oDS_PH_PY006B.setValue("U_UPSTDAMT", i, oRecordSet.Fields(15).VALUE)
        Call oDS_PH_PY006B.setValue("U_UPBNSAMT", i, oRecordSet.Fields(16).VALUE)
        Call oDS_PH_PY006B.setValue("U_APPLYYN", i, oRecordSet.Fields(17).VALUE)
        Call oDS_PH_PY006B.setValue("U_LineMemo", i, oRecordSet.Fields(18).VALUE)
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
PH_PY006_MTX01_Exit:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Exit Sub
PH_PY006_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY006_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY006_MTX02()

    '//승호처리 및 발령사항 추가
    
    Dim i       As Long
    Dim sQry    As String
    
    Dim Param01 As String
    Dim Param02 As String
    Dim Param03 As String
    
    Dim CLTCOD As String
    Dim DocDate As String
    Dim appNum As String
    Dim JIGTYP As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY006_MTX02_Error

    Call oForm.Freeze(True)
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    
    Param01 = oForm.Items("DocEntry").Specific.VALUE
    Param02 = oForm.Items("appNum").Specific.VALUE
    Param03 = oForm.Items("Canceled").Specific.VALUE
    
    CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE)
    DocDate = Trim(oForm.Items("DocDate").Specific.VALUE)
    appNum = oForm.Items("appNum").Specific.VALUE
    JIGTYP = oForm.Items("JIGTYP").Specific.VALUE
    
    sQry = "Select Count(*) From [@PH_PY006A] a Inner Join [@PH_PY006B] b On a.DocEntry = b.DocEntry "
    sQry = sQry + " Inner Join [@PH_PY001G] c On b.U_MSTCOD = c.Code "
    sQry = sQry + " Where a.DocEntry = '" & Param01 & "' and c.U_appNum = a.U_appNum and c.U_appType = '08' "
    sQry = sQry + " and c.U_appDate = '" & DocDate & "'"
    
    Call oRecordSet.DoQuery(sQry)
    If (oRecordSet.Fields(0).VALUE > 0) Then
        Call MDC_Com.MDC_GF_Message("승호작업을 이미 처리했습니다. 확인바랍니다.", "E")
        GoTo PH_PY006_MTX02_Exit
    End If
    
    sQry = "EXEC [PH_PY006_02] '" & Param01 & "', '" & Param02 & "', '" & Param03 & "'"
    Call oRecordSet.DoQuery(sQry)
    
    Call MDC_Com.MDC_GF_Message("승호작업을 정상 처리했습니다.", "S")
    
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Sub
PH_PY006_MTX02_Exit:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    
    Exit Sub
PH_PY006_MTX02_Error:
    
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY006_MTX02_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub
Private Sub PH_PY006_MTX03()

    '//승호취소 및 발령관리 삭제
    
    Dim i       As Long
    Dim sQry    As String
    Dim DocDate As String
    
    Dim Param01 As String
    Dim Param02 As String
    Dim Param03 As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY006_MTX02_Error

    Call oForm.Freeze(True)
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Param01 = oForm.Items("DocEntry").Specific.VALUE
    Param02 = oForm.Items("appNum").Specific.VALUE
    Param03 = oForm.Items("Canceled").Specific.VALUE
    
    DocDate = Trim(oForm.Items("DocDate").Specific.VALUE)
        
    sQry = "Select Count(*) From [@PH_PY006A] a Inner Join [@PH_PY006B] b On a.DocEntry = b.DocEntry "
    sQry = sQry + " Inner Join [@PH_PY001G] c On b.U_MSTCOD = c.Code "
    sQry = sQry + " Where a.DocEntry = '" & Param01 & "' and c.U_appNum = a.U_appNum and c.U_appType = '08' "
    sQry = sQry + " and c.U_appDate = '" & DocDate & "'"
    
  
    Call oRecordSet.DoQuery(sQry)
    If (oRecordSet.Fields(0).VALUE <= 0) Then
        Call MDC_Com.MDC_GF_Message("승호취소 할 자료가 대상이 없거나 이미 취소했습니다. 확인바랍니다.", "E")
        GoTo PH_PY006_MTX02_Exit
    End If
    
    
    
    sQry = "EXEC [PH_PY006_02] '" & Param01 & "', '" & Param02 & "', '" & Param03 & "'"
    Call oRecordSet.DoQuery(sQry)
    
    Call MDC_Com.MDC_GF_Message("승호작업을 취소 처리했습니다.", "S")
    
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Sub
PH_PY006_MTX02_Exit:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    
    Exit Sub
PH_PY006_MTX02_Error:
    
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY006_MTX02_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY006_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY006_Validate_Error
    PH_PY006_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY006A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", bmt_Short, True
        PH_PY006_Validate = False
        GoTo PH_PY006_Validate_Exit
    End If
'
    If ValidateType = "수정" Then

    ElseIf ValidateType = "행삭제" Then

    ElseIf ValidateType = "취소" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY006_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY006_Validate_Error:
    PH_PY006_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY006_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function
