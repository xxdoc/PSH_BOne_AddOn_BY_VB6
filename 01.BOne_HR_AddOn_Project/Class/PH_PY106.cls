VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY106"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY106.cls
'//  Module         : 인사관리 > 급여관리
'//  Desc           : 수당계산식설정
'//********************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form


''// 매트릭스 사용시
Public oMat1 As SAPbouiCOM.Matrix
Public oMat2 As SAPbouiCOM.Matrix
Public oMat3 As SAPbouiCOM.Matrix

Private oDS_PH_PY106A As SAPbouiCOM.DBDataSource
Private oDS_PH_PY106B As SAPbouiCOM.DBDataSource
Private oDS_PH_PY106C As SAPbouiCOM.DBDataSource
Private oDS_PH_PY106D As SAPbouiCOM.DBDataSource

Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long

Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY106.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY106_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY106"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)


    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
    oForm.DataBrowser.BrowseBy = "Code"
    oForm.PaneLevel = 1
    
    oForm.Freeze True
    Call PH_PY106_CreateItems
    Call PH_PY106_EnableMenus
    Call PH_PY106_SetDocument(oFromDocEntry01)
'    Call PH_PY106_FormResize
    
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

Private Function PH_PY106_CreateItems() As Boolean

    Dim sQry        As String
    Dim i           As Long
    
    Dim oCheck      As SAPbouiCOM.CheckBox
    Dim oEdit       As SAPbouiCOM.EditText
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns    As SAPbouiCOM.Columns
    Dim optBtn      As SAPbouiCOM.OptionBtn
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY106_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
'    '//매트릭스--------------------------------------------------------------------------------------
    Set oDS_PH_PY106A = oForm.DataSources.DBDataSources("@PH_PY106A")   '//헤더
    Set oDS_PH_PY106B = oForm.DataSources.DBDataSources("@PH_PY106B")   '//라인
    Set oDS_PH_PY106C = oForm.DataSources.DBDataSources("@PH_PY106C")   '//라인
    Set oDS_PH_PY106D = oForm.DataSources.DBDataSources("@PH_PY106D")   '//라인

    Call oForm.DataSources.UserDataSources.Add("DISSIL", dt_LONG_TEXT)
    Set oEdit = oForm.Items("DISSIL").Specific      '//공식
    oEdit.DataBind.SetBound True, "", "DISSIL"
    
    Set oMat1 = oForm.Items("Mat1").Specific
    Set oMat2 = oForm.Items("Mat2").Specific
    Set oMat3 = oForm.Items("Mat3").Specific

    oMat1.SelectionMode = ms_NotSupported
    oMat1.AutoResizeColumns
    oMat3.SelectionMode = ms_NotSupported
    oMat3.AutoResizeColumns

    '//----------------------------------------------------------------------------------------------
    '// 헤더 설정
    '//----------------------------------------------------------------------------------------------
     '/ 귀속년월
     oForm.Items("YM").Specific.Value = Format$(Now, "YYYYMM")
     
     '//사업장
    Set oCombo = oForm.Items("CLTCOD").Specific
'    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
'    Call SetReDataCombo(oForm, sQry, oCombo)
    oForm.Items("CLTCOD").DisplayDesc = True
    
    '// 1.급여형태-계산식
    Set oCombo = oForm.Items("PAYTYP").Specific
    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P132' AND U_UseYN= 'Y'"
    Call SetReDataCombo(oForm, sQry, oCombo)
    oForm.Items("PAYTYP").DisplayDesc = True
    
    '// 1.근속년수 계산기준
    Set oCombo = oForm.Items("GNSGBN").Specific
    oCombo.ValidValues.Add "1", "그룹입사일"
    oCombo.ValidValues.Add "2", "입사  일자"
    If oCombo.ValidValues.Count >= 1 Then
        oCombo.Select 1, psk_Index
    End If
    
    '// 2.상여 계산단위 (
    Set oCombo = oForm.Items("BNSLEN").Specific
    oCombo.ValidValues.Add "1", "  원"
    oCombo.ValidValues.Add "10", "십원"
    oCombo.ValidValues.Add "100", "백원"
    oCombo.ValidValues.Add "1000", "천원"
    If oCombo.ValidValues.Count >= 1 Then
        oCombo.Select 0, psk_Index
    End If
    
    '// 3.상여 끝전처리
    Set oCombo = oForm.Items("BNSRND").Specific
    oCombo.ValidValues.Add "R", "반올림"
    oCombo.ValidValues.Add "F", "절사"
    oCombo.ValidValues.Add "C", "절상"
    If oCombo.ValidValues.Count >= 1 Then
        oCombo.Select 0, psk_Index
    End If
    
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
    
PH_PY106_CreateItems_Error:
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY106_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY106_EnableMenus()
    
    On Error GoTo PH_PY106_EnableMenus_Error

    Call oForm.EnableMenu("1283", True)     '//제거
    Call oForm.EnableMenu("1284", False)    '//취소
    Call oForm.EnableMenu("1293", True)     '//행삭제
    
    Exit Sub
    
PH_PY106_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY106_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY106_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY106_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY106_FormItemEnabled
        Call PH_PY106_AddMatrixRow
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY106_FormItemEnabled
        oForm.Items("Code").Specific.Value = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY106_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY106_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY106_FormItemEnabled()
    Dim oCombo      As SAPbouiCOM.ComboBox
    
    On Error GoTo PH_PY106_FormItemEnabled_Error
    
    
    Call oForm.Freeze(True)
    If (oForm.Mode = fm_ADD_MODE) Then
        oForm.Items("CLTCOD").Enabled = True
        oForm.Items("YM").Enabled = True
        oForm.Items("PAYTYP").Enabled = True
        oForm.Items("GNSGBN").Enabled = True
        oForm.Items("BNSLEN").Enabled = True
        oForm.Items("BNSRND").Enabled = True
'        oMat1.Columns("CSUCOD").Editable = False
'        oMat1.Columns("CSUNAM").Editable = False
        
        Call PH_PY106_Display_CsuItem                    '//Mat2 초기값 가져옴
        
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        
        '/ 귀속년월
        oForm.Items("YM").Specific.Value = Format$(Now, "YYYYMM")
        
        '// 1.근속년수 계산기준
        Set oCombo = oForm.Items("GNSGBN").Specific
        oCombo.Select 1, psk_Index
        
        '// 2.상여 계산단위 (
        Set oCombo = oForm.Items("BNSLEN").Specific
        oCombo.Select 0, psk_Index
        
        '// 3.상여 끝전처리
        Set oCombo = oForm.Items("BNSRND").Specific
        oCombo.Select 0, psk_Index

        Call oForm.EnableMenu("1293", True)     '//행삭제
        Call oForm.EnableMenu("1283", True)     '//제거
        
    ElseIf (oForm.Mode = fm_FIND_MODE) Then
        oForm.Items("CLTCOD").Enabled = True
        oForm.Items("YM").Enabled = True
        oForm.Items("PAYTYP").Enabled = True
        oForm.Items("GNSGBN").Enabled = True
        oForm.Items("BNSLEN").Enabled = True
        oForm.Items("BNSRND").Enabled = True
'        oMat1.Columns("CSUCOD").Editable = False
'        oMat1.Columns("CSUNAM").Editable = False
        
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        
        Call oForm.EnableMenu("1293", True)     '//행삭제
        Call oForm.EnableMenu("1283", True)     '//제거
    ElseIf (oForm.Mode = fm_OK_MODE) Then
        oForm.Items("CLTCOD").Enabled = False
        oForm.Items("YM").Enabled = False
        oForm.Items("PAYTYP").Enabled = False
        oForm.Items("GNSGBN").Enabled = False
        oForm.Items("BNSLEN").Enabled = False
        oForm.Items("BNSRND").Enabled = False
'        oMat1.Columns("CSUCOD").Editable = False
'        oMat1.Columns("CSUNAM").Editable = False
        
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD", False)
        
        Call oForm.EnableMenu("1293", True)     '//행삭제
        Call oForm.EnableMenu("1283", True)     '//제거

    End If
    Call oForm.Freeze(False)
    Exit Sub
    
PH_PY106_FormItemEnabled_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY106_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
                    If oForm.Mode = fm_ADD_MODE Or oForm.Mode = fm_UPDATE_MODE Then
                        If PH_PY106_DataValidCheck = False Then
                            BubbleEvent = False
                            Exit Sub
                        Else
                           If MatrixSpaceLineDel = False Then
                               BubbleEvent = False
                           End If
                        End If
                    End If
                End If
            ElseIf pval.BeforeAction = False Then
                If pval.ItemUid = "FLD01" Or pval.ItemUid = "FLD02" Then
                    oForm.PaneLevel = Right(pval.ItemUid, 2)
                End If
                If pval.ItemUid = "1" And pval.ActionSuccess = True And oForm.Mode = fm_ADD_MODE Then
                    oForm.Mode = fm_OK_MODE
                    Call Sbo_Application.ActivateMenuItem("1282")   '/문서추가
                ElseIf pval.ItemUid = "Btn1" Then  '/ 가져오기
                        If PH_PY106_DataValidCheck = False Then
                            BubbleEvent = False
                            Exit Sub
                        Else
                            Call Display_PH_PY102
                        End If
                ElseIf pval.ItemUid = "Btn2" Then '/ 공식 검증
                    Call PH_PY106_Display_CsuItem
                End If
            End If
'----------------------------------------------------------
        Case et_KEY_DOWN: '//2
            'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
            '//추가모드에서 코드이벤트가 코드에서 일어 났을때
            'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
             If pval.BeforeAction = True And pval.ItemUid = "YM" And pval.CharPressed = 9 And pval.FormMode <> fm_FIND_MODE Then
                If oMat1.RowCount > 0 Then
                    oMat1.Columns("LINSEQ").Cells(oMat1.VisualRowCount).CLICK ct_Regular
                    BubbleEvent = False
                End If
             ElseIf pval.BeforeAction = True And pval.ColUID = "LINSEQ" And pval.CharPressed = 9 Then
                If Trim$(oMat1.Columns(pval.ColUID).Cells(pval.Row).Specific.String) = "" Then
                    Call Sbo_Application.StatusBar.SetText("순서는 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error)
                    BubbleEvent = False
                End If
             End If
'----------------------------------------------------------
        Case et_GOT_FOCUS: '//3
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
            If pval.FormUID = oForm.uniqueID And pval.BeforeAction = True And oLastItemUID = "Mat1" And oLastColUID = "LINSEQ" And oLastColRow > 0 _
                And (oLastItemUID <> pval.ColUID Or oLastColRow <> pval.Row) And pval.ItemUid <> "1000001" And pval.ItemUid <> "2" Then
                If oLastColRow > oMat1.VisualRowCount Then
                    Exit Sub
                End If
            ElseIf pval.FormMode <> fm_FIND_MODE And pval.BeforeAction = False Then
                If pval.ItemUid = "Mat1" And pval.Row > 0 Then
                   oForm.DataSources.UserDataSources("DISSIL").ValueEx = oMat1.Columns("SILCUN").Cells(pval.Row).Specific.Value
                ElseIf pval.ItemUid = "Mat2" And pval.Row > 0 Then
                    oForm.DataSources.UserDataSources("DISSIL").ValueEx = oForm.DataSources.UserDataSources("DISSIL").ValueEx & oMat2.Columns("FILCOD").Cells(pval.Row).Specific.Value
                End If
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
                    If pval.ItemUid = "Mat1" Then
                        If pval.ColUID = "CSUCOD" Then
                            Call PH_PY106_AddMatrixRow
                            oMat1.Columns(pval.ColUID).Cells(pval.Row).CLICK ct_Regular
                        End If
                    End If
                    If pval.ItemUid = "Mat3" Then
                        If pval.ColUID = "CSUCOD" Then
                            Call PH_PY106_AddMatrixRow
                            oMat3.Columns(pval.ColUID).Cells(pval.Row).CLICK ct_Regular
                        End If
                    End If
                End If
            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_MATRIX_LOAD: '//11
            If pval.BeforeAction = True Then
            
            ElseIf pval.BeforeAction = False Then
                oMat1.LoadFromDataSource
                oMat2.LoadFromDataSource
                oMat3.LoadFromDataSource
                Call PH_PY106_FormItemEnabled
                Call PH_PY106_AddMatrixRow
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
                Set oDS_PH_PY106A = Nothing
                Set oDS_PH_PY106B = Nothing
                Set oDS_PH_PY106C = Nothing
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
            If pval.BeforeAction = True Then
        
            ElseIf pval.BeforeAction = False Then
                oForm.Freeze True
                oForm.Items("Mat2").Top = oForm.Items("Mat1").Top
                oForm.Items("Mat2").Left = oForm.Items("Mat1").Width + 15
                oForm.Items("Mat2").Width = "240"
                oForm.Items("Mat2").Height = oForm.Items("Mat1").Height
                oMat2.Columns("Code").Width = "20"
                oMat2.Columns("Name").Width = "30"
                oMat2.Columns("FILCOD").Width = "90"
                oMat2.Columns("REMARK").Width = "80"
                oForm.Freeze False
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
                Call AuthorityCheck(oForm, "CLTCOD", "@PH_PY106A", "Code")      '//접속자 권한에 따른 사업장 보기
        End Select
    ElseIf (pval.BeforeAction = False) Then
        Select Case pval.MenuUID
            Case "1283":
                oForm.Mode = fm_ADD_MODE
                Call PH_PY106_FormItemEnabled
                Call PH_PY106_AddMatrixRow
            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//문서찾기
                Call PH_PY106_FormItemEnabled
                Call PH_PY106_AddMatrixRow
            Case "1282":    '//문서추가
                Call PH_PY106_FormItemEnabled
                Call PH_PY106_AddMatrixRow
            Case "1288", "1289", "1290", "1291":
                Call PH_PY106_FormItemEnabled
            Case "1293" '// 행삭제
                '// [MAT1 용]
                 If oMat1.RowCount <> oMat1.VisualRowCount Then
                    oMat1.FlushToDataSource

                    While (i <= oDS_PH_PY106B.Size - 1)
                        If oDS_PH_PY106B.GetValue("U_FILD01", i) = "" Then
                            oDS_PH_PY106B.RemoveRecord (i)
                            i = 0
                        Else
                            i = i + 1
                        End If
                    Wend

                    For i = 0 To oDS_PH_PY106B.Size
                        Call oDS_PH_PY106B.setValue("U_LineNum", i, i + 1)
                    Next i

                    oMat1.LoadFromDataSource
                End If
'                '// [MAT2 용]
'                 If oMat2.RowCount <> oMat2.VisualRowCount Then
'                    oMat2.FlushToDataSource
'
'                    While (i <= oDS_PH_PY106C.Size - 1)
'                        If oDS_PH_PY106C.GetValue("U_FILD01", i) = "" Then
'                            oDS_PH_PY106C.RemoveRecord (i)
'                            i = 0
'                        Else
'                            i = i + 1
'                        End If
'                    Wend
'
'                    For i = 0 To oDS_PH_PY106B.Size
'                        Call oDS_PH_PY106B.setValue("U_LineNum", i, i + 1)
'                    Next i
'
'                    oMat1.LoadFromDataSource
'                End If
                Call PH_PY106_AddMatrixRow
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

Sub PH_PY106_AddMatrixRow()
    Dim oRow    As Long
    
    On Error GoTo PH_PY106_AddMatrixRow_Error
    
    Call oForm.Freeze(True)
    
    '//[Mat1 용]
    oMat1.FlushToDataSource
    oRow = oMat1.VisualRowCount
'
    If oMat1.VisualRowCount > 0 Then
        If Trim(oDS_PH_PY106B.GetValue("U_CSUCOD", oRow - 1)) <> "" Then
            If oDS_PH_PY106B.Size <= oMat1.VisualRowCount Then
                oDS_PH_PY106B.InsertRecord (oRow)
            End If
            oDS_PH_PY106B.Offset = oRow
            oDS_PH_PY106B.setValue "U_LineNum", oRow, oRow + 1
            oDS_PH_PY106B.setValue "U_LINSEQ", oRow, ""
            oDS_PH_PY106B.setValue "U_CSUCOD", oRow, ""
            oDS_PH_PY106B.setValue "U_CSUNAM", oRow, ""
            oDS_PH_PY106B.setValue "U_SILCUN", oRow, ""
            oDS_PH_PY106B.setValue "U_SILCOD", oRow, ""
            oDS_PH_PY106B.setValue "U_BNSBAS", oRow, "N"
            oDS_PH_PY106B.setValue "U_REMARK", oRow, ""
            oMat1.LoadFromDataSource
        Else
            oDS_PH_PY106B.Offset = oRow - 1
            oDS_PH_PY106B.setValue "U_LineNum", oRow - 1, oRow
            oDS_PH_PY106B.setValue "U_LINSEQ", oRow - 1, ""
            oDS_PH_PY106B.setValue "U_CSUCOD", oRow - 1, ""
            oDS_PH_PY106B.setValue "U_CSUNAM", oRow - 1, ""
            oDS_PH_PY106B.setValue "U_SILCUN", oRow - 1, ""
            oDS_PH_PY106B.setValue "U_SILCOD", oRow - 1, ""
            oDS_PH_PY106B.setValue "U_BNSBAS", oRow - 1, "N"
            oDS_PH_PY106B.setValue "U_REMARK", oRow - 1, ""
            oMat1.LoadFromDataSource
        End If
    ElseIf oMat1.VisualRowCount = 0 Then
        oDS_PH_PY106B.Offset = oRow
        oDS_PH_PY106B.setValue "U_LineNum", oRow, oRow + 1
        oDS_PH_PY106B.setValue "U_LINSEQ", oRow, ""
        oDS_PH_PY106B.setValue "U_CSUCOD", oRow, ""
        oDS_PH_PY106B.setValue "U_CSUNAM", oRow, ""
        oDS_PH_PY106B.setValue "U_SILCUN", oRow, ""
        oDS_PH_PY106B.setValue "U_SILCOD", oRow, ""
        oDS_PH_PY106B.setValue "U_BNSBAS", oRow, "N"
        oDS_PH_PY106B.setValue "U_REMARK", oRow, ""
        oMat1.LoadFromDataSource
    End If
'
    '//[Mat2 용]
    oMat3.FlushToDataSource
    oRow = oMat3.VisualRowCount
'
    If oMat3.VisualRowCount > 0 Then
        If Trim(oDS_PH_PY106D.GetValue("U_CSUCOD", oRow - 1)) <> "" Then
            If oDS_PH_PY106D.Size <= oMat3.VisualRowCount Then
                oDS_PH_PY106D.InsertRecord (oRow)
            End If
            oDS_PH_PY106D.Offset = oRow
            oDS_PH_PY106D.setValue "U_LineNum", oRow, oRow + 1
            oDS_PH_PY106D.setValue "U_LINSEQ", oRow, ""
            oDS_PH_PY106D.setValue "U_Status", oRow, ""
            oDS_PH_PY106D.setValue "U_WorkType", oRow, ""
            oDS_PH_PY106D.setValue "U_Order", oRow, ""
            oDS_PH_PY106D.setValue "U_CSUCOD", oRow, ""
            oDS_PH_PY106D.setValue "U_SILCUN", oRow, ""
            oDS_PH_PY106D.setValue "U_REMARK", oRow, ""
            oMat3.LoadFromDataSource
        Else
            oDS_PH_PY106D.Offset = oRow - 1
            oDS_PH_PY106D.setValue "U_LineNum", oRow - 1, oRow
            oDS_PH_PY106D.setValue "U_LINSEQ", oRow - 1, ""
            oDS_PH_PY106D.setValue "U_Status", oRow - 1, ""
            oDS_PH_PY106D.setValue "U_WorkType", oRow - 1, ""
            oDS_PH_PY106D.setValue "U_Order", oRow - 1, ""
            oDS_PH_PY106D.setValue "U_CSUCOD", oRow - 1, ""
            oDS_PH_PY106D.setValue "U_SILCUN", oRow - 1, ""
            oDS_PH_PY106D.setValue "U_REMARK", oRow - 1, ""
            oMat3.LoadFromDataSource
        End If
    ElseIf oMat3.VisualRowCount = 0 Then
        oDS_PH_PY106D.Offset = oRow
        oDS_PH_PY106D.setValue "U_LineNum", oRow, oRow + 1
        oDS_PH_PY106D.setValue "U_LINSEQ", oRow, ""
        oDS_PH_PY106D.setValue "U_Status", oRow, ""
        oDS_PH_PY106D.setValue "U_WorkType", oRow, ""
        oDS_PH_PY106D.setValue "U_Order", oRow, ""
        oDS_PH_PY106D.setValue "U_CSUCOD", oRow, ""
        oDS_PH_PY106D.setValue "U_SILCUN", oRow, ""
        oDS_PH_PY106D.setValue "U_REMARK", oRow, ""
        oMat3.LoadFromDataSource
    End If
'
    Call oForm.Freeze(False)
    Exit Sub
PH_PY106_AddMatrixRow_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY106_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY106_FormClear()
On Error GoTo PH_PY106_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY106'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.Value = 1
    Else
        oForm.Items("DocEntry").Specific.Value = DocEntry
    End If
    Exit Sub
PH_PY106_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY106_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY106_DataValidCheck() As Boolean
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim ExistYN          As String
        
    On Error GoTo PH_PY106_DataValidCheck_Error
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    PH_PY106_DataValidCheck = False
    
    
 '/ Check
    Select Case True
    Case oDS_PH_PY106A.GetValue("U_CLTCOD", 0) = ""
        Sbo_Application.StatusBar.SetText "사업장은 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
        oForm.Items("CLTCOD").CLICK ct_Regular
        PH_PY106_DataValidCheck = False
        Exit Function
    Case MDC_SetMod.ChkYearMonth(oDS_PH_PY106A.GetValue("U_YM", 0)) = False
        Sbo_Application.StatusBar.SetText "작업연월은 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
        oForm.Items("YM").CLICK ct_Regular
        PH_PY106_DataValidCheck = False
        Exit Function
    Case oDS_PH_PY106A.GetValue("U_PAYTYP", 0) = ""
        Sbo_Application.StatusBar.SetText "급여형태는 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
        oForm.Items("PAYTYP").CLICK ct_Regular
        PH_PY106_DataValidCheck = False
        Exit Function
    Case oDS_PH_PY106A.GetValue("U_GNSGBN", 0) = ""
        Sbo_Application.StatusBar.SetText "근속일계산기준은 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
        oForm.Items("GNSGBN").CLICK ct_Regular
        PH_PY106_DataValidCheck = False
        Exit Function
    Case oDS_PH_PY106A.GetValue("U_BNSLEN", 0) = ""
        Sbo_Application.StatusBar.SetText "상여 계산단위는 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
        oForm.Items("BNSLEN").CLICK ct_Regular
        PH_PY106_DataValidCheck = False
        Exit Function
    Case oDS_PH_PY106A.GetValue("U_BNSRND", 0) = ""
        Sbo_Application.StatusBar.SetText "상여 끝전처리는 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
        oForm.Items("BNSRND").CLICK ct_Regular
        PH_PY106_DataValidCheck = False
        Exit Function
    End Select
    
    '// Code & Name 생성
    Call oDS_PH_PY106A.setValue("Code", 0, Trim$(oDS_PH_PY106A.GetValue("U_CLTCOD", 0)) & Trim$(oDS_PH_PY106A.GetValue("U_YM", 0)) & Trim$(oDS_PH_PY106A.GetValue("U_PAYTYP", 0)))
    Call oDS_PH_PY106A.setValue("NAME", 0, Trim$(oDS_PH_PY106A.GetValue("U_CLTCOD", 0)) & Trim$(oDS_PH_PY106A.GetValue("U_YM", 0)) & Trim$(oDS_PH_PY106A.GetValue("U_PAYTYP", 0)))
    
    If oForm.Mode = fm_ADD_MODE Then
        '//저장된 데이터 체크
        sQry = "SELECT Top 1 Code FROM [@PH_PY106A] "
        sQry = sQry & " WHERE Code = '" & Trim(oDS_PH_PY106A.GetValue("Code", 0)) & "'"
        oRecordSet.DoQuery sQry
        
        If oRecordSet.Fields.Count > 0 Then
            ExistYN = Trim$(oRecordSet.Fields(0).Value)
        End If
        
        If ExistYN <> "" And Trim$(oDS_PH_PY106A.GetValue("Code", 0)) <> ExistYN Then
            Sbo_Application.StatusBar.SetText "Code" & "데이터가 일치합니다. 저장되지 않습니다. 확인하여 주십시오.", bmt_Short, smt_Error
        End If
    End If

    PH_PY106_DataValidCheck = True
    Set oRecordSet = Nothing
    Exit Function
    
PH_PY106_DataValidCheck_Error:
    Set oRecordSet = Nothing
    PH_PY106_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY106_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Function MatrixSpaceLineDel() As Boolean
'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
'저장할 데이터의 유효성을 점검한다
'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
On Error GoTo Error_Message
    Dim iRow        As Long
    Dim kRow        As Long
    Dim ErrNum      As Integer
    Dim Chk_Data    As String
    
    ErrNum = 0
    
    oMat1.FlushToDataSource
        
    If oMat1.RowCount = 1 Then
        ErrNum = 1
        GoTo Error_Message
    End If
    
    For iRow = 0 To oMat1.VisualRowCount - 2
        oDS_PH_PY106B.Offset = iRow
        If Trim$(oDS_PH_PY106B.GetValue("U_CSUCOD", iRow)) = "" Then
            ErrNum = 2
            oMat1.Columns("CSUCOD").Cells(iRow + 1).CLICK ct_Regular
            GoTo Error_Message
        ElseIf Trim$(oDS_PH_PY106B.GetValue("U_LINSEQ", iRow)) = "" And Left$(oDS_PH_PY106B.GetValue("U_CSUCOD", iRow), 1) <> "X" Then
            ErrNum = 5
            oMat1.Columns("LINSEQ").Cells(iRow + 1).CLICK ct_Regular
            GoTo Error_Message
        ElseIf Trim$(oDS_PH_PY106B.GetValue("U_SILCUN", iRow)) = "" Then
            ErrNum = 4
            oMat1.Columns("SILCUN").Cells(iRow + 1).CLICK ct_Regular
            GoTo Error_Message
        ElseIf Trim$(oDS_PH_PY106B.GetValue("U_BNSBAS", iRow)) = "Y" Then
            If Left$(Trim$(oDS_PH_PY106B.GetValue("U_CSUCOD", iRow)), 1) = "X" Then
                ErrNum = 6
                oMat1.Columns("BNSBAS").Cells(iRow + 1).CLICK ct_Regular
                GoTo Error_Message
            ElseIf Trim$(oDS_PH_PY106B.GetValue("U_CSUCOD", iRow)) = "A04" Then
                ErrNum = 7
                oMat1.Columns("CSUCOD").Cells(iRow + 1).CLICK ct_Regular
                GoTo Error_Message
            
            End If
        Else
            'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
            '중복체크작업
            'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
            Chk_Data = Trim$(oDS_PH_PY106B.GetValue("U_CSUCOD", iRow))
            For kRow = iRow + 1 To oMat1.VisualRowCount - 2
                oDS_PH_PY106B.Offset = kRow
                If Trim$(Chk_Data) = Trim$(oDS_PH_PY106B.GetValue("U_CSUCOD", kRow)) Then
                    ErrNum = 3
                    oMat1.Columns("LINSEQ").Cells(iRow + 1).CLICK ct_Regular
                    GoTo Error_Message
                End If
            Next kRow
        End If
    Next iRow
    
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '//맨마지막에 데이터를 삭제하는 이유는 행을 추가 할경우에 디비데이터소스에
    '//이미 데이터가 들어가 있기 때문에 저장시에는 마지막 행(DB데이터 소스에)을 삭제한다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    oDS_PH_PY106B.RemoveRecord oDS_PH_PY106B.Size - 1   '// Mat1에 마지막라인(빈라인) 삭제
    
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '행을 삭제하였으니 DB데이터 소스를 다시 가져온다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    oMat1.LoadFromDataSource
    
    MatrixSpaceLineDel = True
    Exit Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Error_Message:
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "입력할 데이터가 없습니다. 입력하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 2 Then
        Sbo_Application.StatusBar.SetText "코드는 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 3 Then
        Sbo_Application.StatusBar.SetText "내용이 중복입력되었습니다. 확인하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 4 Then
        Sbo_Application.StatusBar.SetText "계산식은 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 5 Then
        Sbo_Application.StatusBar.SetText "순서는 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 6 Then
        Sbo_Application.StatusBar.SetText "기본일급/통상일급/기본시급/통상시급은 상여지정을 할 수 없습니다. 확인하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 7 Then
        Sbo_Application.StatusBar.SetText "상여금에는 상여지정을 할 수 없습니다. 확인하여 주십시오.", bmt_Short, smt_Error
    Else
        Sbo_Application.StatusBar.SetText "MatrixSpaceLineDel Error :" & Space$(10) & Err.Description, bmt_Short, smt_Error
    End If
    MatrixSpaceLineDel = False
End Function

Function PH_PY106_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY106_Validate_Error
    PH_PY106_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY106A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.Value & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", bmt_Short, True
        PH_PY106_Validate = False
        GoTo PH_PY106_Validate_Exit
    End If
'
    If ValidateType = "수정" Then

    ElseIf ValidateType = "행삭제" Then

    ElseIf ValidateType = "취소" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY106_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY106_Validate_Error:
    PH_PY106_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY106_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function



Private Sub Display_PH_PY102()  '/수당
On Error GoTo Error_Message
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim oCLTCOD         As String
    Dim oJOBYMM         As String
    Dim i               As Integer
    Dim cnt             As Integer
    Dim CSUCOD          As String
    Dim SILCUN          As String
    Dim SILTYP          As String
    
'/ Matrix2 초기화
    cnt = oDS_PH_PY106B.Size
    If cnt > 0 Then
        For i = 1 To cnt - 1
            oDS_PH_PY106B.RemoveRecord oDS_PH_PY106B.Size - 1
        Next i
    Else
        oMat1.LoadFromDataSource
    End If
    oCLTCOD = Trim$(oDS_PH_PY106A.GetValue("U_CLTCOD", 0))
    oJOBYMM = Trim$(oDS_PH_PY106A.GetValue("U_YM", 0))
    i = 0
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
'/ 기본셋팅값 가져오기
    sQry = "SELECT Code, Name, U_FILCOD ,U_REMARK FROM [@PH_PY106C] WHERE Code BETWEEN 'X01' AND 'X05' ORDER BY CODE"
    oRecordSet.DoQuery sQry
    Do Until oRecordSet.EOF
        If i + 1 > oDS_PH_PY106B.Size Then
            oDS_PH_PY106B.InsertRecord (i)
        End If
        oDS_PH_PY106B.Offset = i
        oDS_PH_PY106B.setValue "U_LineNum", i, i + 1
        oDS_PH_PY106B.setValue "U_LINSEQ", i, ""
        oDS_PH_PY106B.setValue "U_CSUCOD", i, Trim(oRecordSet.Fields(0).Value)
        oDS_PH_PY106B.setValue "U_CSUNAM", i, Trim(oRecordSet.Fields(1).Value)
        oDS_PH_PY106B.setValue "U_SILCOD", i, ""
        oDS_PH_PY106B.setValue "U_SILCUN", i, ""
        oDS_PH_PY106B.setValue "U_BNSBAS", i, "N"
        oDS_PH_PY106B.setValue "U_REMARK", i, Trim(oRecordSet.Fields(3).Value)
        oRecordSet.MoveNext
        i = i + 1
    Loop
    cnt = i
    sQry = "Exec PH_PY102 '" & Trim$(oCLTCOD) & "','" & Trim$(oJOBYMM) & "', '', '', '', ''"
    oRecordSet.DoQuery sQry
    Do Until oRecordSet.EOF
        CSUCOD = ""
        SILTYP = ""
        If i + 1 > oDS_PH_PY106B.Size Then
            oDS_PH_PY106B.InsertRecord (i)
        End If
        CSUCOD = Trim(oRecordSet.Fields("U_CSUCOD").Value)
        oDS_PH_PY106B.Offset = i
        oDS_PH_PY106B.setValue "U_LineNum", i, i + 1
        oDS_PH_PY106B.setValue "U_LINSEQ", i, i - cnt + 1
        oDS_PH_PY106B.setValue "U_CSUCOD", i, CSUCOD
        oDS_PH_PY106B.setValue "U_CSUNAM", i, Trim(oRecordSet.Fields("U_CSUNAM").Value)
        oDS_PH_PY106B.setValue "U_SILCOD", i, ""
        oDS_PH_PY106B.setValue "U_BNSBAS", i, "N"
        oDS_PH_PY106B.setValue "U_REMARK", i, ""
        If Trim$(CSUCOD) = "A01" Then
            SILCUN = "T1.U_STDAMT"
        Else
            SILTYP = MDC_SetMod.Get_ReData("U_FIXGBN + U_INSLIN", "U_CSUCOD", "[@PH_PY102B]", "'" & CSUCOD & "'", " AND Code = '" & Trim$(oCLTCOD) + Trim$(oJOBYMM) & "'")
            If Left$(SILTYP, 1) = "Y" Then
                SILCUN = "T2.U_CSUD" & Format$(Mid$(SILTYP, 2), "00")
            Else
                SILCUN = "0"
            End If
        End If
        
        oDS_PH_PY106B.setValue "U_SILCUN", i, SILCUN
        
        i = i + 1
        oRecordSet.MoveNext
    Loop
    oMat1.LoadFromDataSource
    PH_PY106_AddMatrixRow
    
    Set oRecordSet = Nothing
    Exit Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Error_Message:
    Set oRecordSet = Nothing
    Sbo_Application.StatusBar.SetText "Display_PH_PY102 Error :" & Space$(10) & Err.Description, bmt_Short, smt_Error
End Sub


Private Sub PH_PY106_Display_CsuItem()
On Error GoTo Error_Message
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim i               As Integer
    Dim cnt             As Integer
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
'/ Matrix2 초기화
    cnt = oDS_PH_PY106C.Size
    If cnt > 0 Then
        For i = 1 To cnt - 1
            oDS_PH_PY106C.RemoveRecord oDS_PH_PY106C.Size - 1
        Next i
    Else
        oMat2.LoadFromDataSource
    End If
    
    i = 0
    
    sQry = "SELECT Code, Name, U_FILCOD, U_REMARK FROM [@PH_PY106C] ORDER BY CODE"
    oRecordSet.DoQuery sQry
    Do Until oRecordSet.EOF
        If i + 1 > oDS_PH_PY106C.Size Then
            oDS_PH_PY106C.InsertRecord (i)
        End If
        oDS_PH_PY106C.Offset = i
        oDS_PH_PY106C.setValue "Code", i, Trim(oRecordSet.Fields(0).Value)
        oDS_PH_PY106C.setValue "Name", i, Trim(oRecordSet.Fields(1).Value)
        oDS_PH_PY106C.setValue "U_FILCOD", i, Trim(oRecordSet.Fields(2).Value)
        oDS_PH_PY106C.setValue "U_REMARK", i, Trim(oRecordSet.Fields(3).Value)
        i = i + 1
        oRecordSet.MoveNext
    Loop
    oMat2.LoadFromDataSource
    oMat2.AutoResizeColumns
    Set oRecordSet = Nothing
    Exit Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Error_Message:
    Set oRecordSet = Nothing
    Sbo_Application.StatusBar.SetText "PH_PY106_Display_CsuItem Error :" & Space$(10) & Err.Description, bmt_Short, smt_Error
End Sub

