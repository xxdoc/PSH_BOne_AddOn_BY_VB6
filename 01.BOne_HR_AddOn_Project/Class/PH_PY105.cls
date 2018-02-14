VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY105"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY105.cls
'//  Module         : 급여관리 > 급여관리
'//  Desc           : 호봉표등록
'//********************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form

Public oMat1 As SAPbouiCOM.Matrix

Private oDS_PH_PY105A As SAPbouiCOM.DBDataSource
Private oDS_PH_PY105B As SAPbouiCOM.DBDataSource

Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long

Private CheckDataApply  As Boolean      '//적용버턴 실행여부
Private CLTCOD          As String       '//사업장
Private YM              As String       '//적용연월

Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY105.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY105_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY105"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
    oForm.DataBrowser.BrowseBy = "Code"
        
    oForm.Freeze True
    Call PH_PY105_CreateItems
    Call PH_PY105_EnableMenus
    Call PH_PY105_SetDocument(oFromDocEntry01)
'    Call PH_PY105_FormResize
    
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

Private Function PH_PY105_CreateItems() As Boolean

    Dim sQry        As String
    Dim i           As Long
    
    Dim oCheck      As SAPbouiCOM.CheckBox
    Dim oEdit       As SAPbouiCOM.EditText
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns    As SAPbouiCOM.Columns
    Dim optBtn      As SAPbouiCOM.OptionBtn
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY105_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Set oDS_PH_PY105A = oForm.DataSources.DBDataSources("@PH_PY105A")
    Set oDS_PH_PY105B = oForm.DataSources.DBDataSources("@PH_PY105B")

    
    Set oMat1 = oForm.Items("Mat1").Specific       '//@PH_PY105B
    
    oMat1.SelectionMode = ms_NotSupported
    oMat1.AutoResizeColumns
    
    CheckDataApply = False
    
    '// 사업장
    Set oCombo = oForm.Items("CLTCOD").Specific
'    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
'    Call SetReDataCombo(oForm, sQry, oCombo)
'    oCombo.Select 0, psk_Index
    oForm.Items("CLTCOD").DisplayDesc = True
    
    
    
     
    '// 직급
    Set oColumn = oMat1.Columns("JIGCOD")
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P129' AND U_UseYN= 'Y'"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        For i = 0 To oRecordSet.RecordCount - 1
            oColumn.ValidValues.Add oRecordSet.Fields(0).VALUE, oRecordSet.Fields(1).VALUE
            oRecordSet.MoveNext
        Next i
    End If
    oColumn.DisplayDesc = True

    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
    
PH_PY105_CreateItems_Error:
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY105_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY105_EnableMenus()
    
    On Error GoTo PH_PY105_EnableMenus_Error

    Call oForm.EnableMenu("1283", True)     '//제거
    Call oForm.EnableMenu("1284", False)    '//취소
    Call oForm.EnableMenu("1293", True)     '//행삭제
    
    Exit Sub
    
PH_PY105_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY105_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY105_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY105_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY105_FormItemEnabled
        Call PH_PY105_AddMatrixRow
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY105_FormItemEnabled
        oForm.Items("Code").Specific.VALUE = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY105_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY105_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY105_FormItemEnabled()
    Dim oCombo      As SAPbouiCOM.ComboBox
    
    On Error GoTo PH_PY105_FormItemEnabled_Error
    
    
    Call oForm.Freeze(True)
    If (oForm.Mode = fm_ADD_MODE) Then
        oForm.Items("CLTCOD").Enabled = True
        oForm.Items("YM").Enabled = True
        oForm.Items("Comments").Enabled = False
        
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        '/ 귀속년월
        oForm.Items("YM").Specific.VALUE = Format$(Now, "YYYYMM")
     
        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", False)    '//문서추가
        
    ElseIf (oForm.Mode = fm_FIND_MODE) Then
        oForm.Items("CLTCOD").Enabled = True
        oForm.Items("YM").Enabled = True
        oForm.Items("Comments").Enabled = False
        
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        
        Call oForm.EnableMenu("1281", False)    '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가
    ElseIf (oForm.Mode = fm_OK_MODE) Then
        oForm.Items("CLTCOD").Enabled = False
        oForm.Items("YM").Enabled = False
        oForm.Items("Comments").Enabled = False
        
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD", False)
        
        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가

    End If
    Call oForm.Freeze(False)
    Exit Sub
    
PH_PY105_FormItemEnabled_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY105_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
                    If oForm.Mode = fm_ADD_MODE Or oForm.Mode = fm_UPDATE_MODE Then
                        If PH_PY105_DataValidCheck = False Then
                            BubbleEvent = False
                        End If
                    End If
                End If
                
            ElseIf pval.BeforeAction = False Then
                If pval.ItemUID = "1" Then
                    If pval.ActionSuccess = True Then
                        If CheckDataApply = True Then
                            Call PH_PY105_DataApply(CLTCOD, YM)
                            CheckDataApply = False
                        End If
                        Call PH_PY105_FormItemEnabled
                        Call PH_PY105_AddMatrixRow
                    End If
                End If
                If pval.ItemUID = "Btn_UPLOAD" Then
                    Call PH_PY105_Excel_Upload
                End If
                If pval.ItemUID = "Btn_Apply" Then
                    CLTCOD = Trim(oDS_PH_PY105A.GetValue("U_CLTCOD", 0))
                    YM = Trim(oDS_PH_PY105A.GetValue("U_YM", 0))
                    If oMat1.RowCount > 1 Then
                        If oForm.Mode = fm_ADD_MODE Or oForm.Mode = fm_UPDATE_MODE Then
                            CheckDataApply = True
                            oForm.Items("1").CLICK ct_Regular
                        ElseIf oForm.Mode = fm_OK_MODE Then
                            Call PH_PY105_DataApply(CLTCOD, YM)
                        End If
                    Else
                        Sbo_Application.SetStatusBarMessage "호봉표 데이터가 없습니다.", bmt_Short, True
                    End If
                End If
            End If
     
'----------------------------------------------------------
        Case et_KEY_DOWN: '//2

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
                    If pval.ItemUID = "Mat1" And pval.ColUID = "HOBCOD" Then
                        Call PH_PY105_AddMatrixRow
                        Call oMat1.Columns(pval.ColUID).Cells(pval.Row).CLICK(ct_Regular)
                    End If
                End If
            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_MATRIX_LOAD: '//11
            If pval.BeforeAction = True Then
            ElseIf pval.BeforeAction = False Then
                oMat1.LoadFromDataSource

                Call PH_PY105_FormItemEnabled
                Call PH_PY105_AddMatrixRow
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
                Set oDS_PH_PY105A = Nothing
                Set oDS_PH_PY105B = Nothing
                
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
'                If pval.ItemUID = "Code" Then
'                    Call MDC_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY105A", "Code")
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
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim i               As Long
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
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
                Call AuthorityCheck(oForm, "CLTCOD", "@PH_PY105A", "Code")      '//접속자 권한에 따른 사업장 보기
                
        End Select
    ElseIf (pval.BeforeAction = False) Then
        Select Case pval.MenuUID
            Case "1283":
                oForm.Mode = fm_ADD_MODE
                Call PH_PY105_FormItemEnabled
                Call PH_PY105_AddMatrixRow
            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//문서찾기
                Call PH_PY105_FormItemEnabled
                Call PH_PY105_AddMatrixRow
                oForm.Items("Code").CLICK ct_Regular
            Case "1282":    '//문서추가
                Call PH_PY105_FormItemEnabled
                Call PH_PY105_AddMatrixRow
            Case "1288", "1289", "1290", "1291":
                Call PH_PY105_FormItemEnabled
            Case "1293" '// 행삭제
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent, oMat1, oDS_PH_PY105B, "U_CODNBR")
                Call PH_PY105_AddMatrixRow
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

Sub PH_PY105_AddMatrixRow()
    Dim oRow    As Long
    
    On Error GoTo PH_PY105_AddMatrixRow_Error
    
    Call oForm.Freeze(True)
    
    '//[Mat1]
    oMat1.FlushToDataSource
    oRow = oMat1.VisualRowCount
    
    If oMat1.VisualRowCount > 0 Then
        If Trim(oDS_PH_PY105B.GetValue("U_HOBCOD", oRow - 1)) <> "" Then
            If oDS_PH_PY105B.Size <= oMat1.VisualRowCount Then
                oDS_PH_PY105B.InsertRecord (oRow)
            End If
            oDS_PH_PY105B.Offset = oRow
            oDS_PH_PY105B.setValue "U_LineNum", oRow, oRow + 1
            oDS_PH_PY105B.setValue "U_JIGCOD", oRow, ""
            oDS_PH_PY105B.setValue "U_HOBCOD", oRow, ""
            oDS_PH_PY105B.setValue "U_HOBNAM", oRow, ""
            oDS_PH_PY105B.setValue "U_STDAMT", oRow, 0
            oDS_PH_PY105B.setValue "U_BNSAMT", oRow, 0
            oDS_PH_PY105B.setValue "U_EXTAMT01", oRow, 0
            oDS_PH_PY105B.setValue "U_EXTAMT02", oRow, 0
            oDS_PH_PY105B.setValue "U_EXTAMT03", oRow, 0
            oDS_PH_PY105B.setValue "U_EXTAMT04", oRow, 0
            oDS_PH_PY105B.setValue "U_EXTAMT05", oRow, 0
            oDS_PH_PY105B.setValue "U_EXTAMT06", oRow, 0
            oDS_PH_PY105B.setValue "U_EXTAMT07", oRow, 0
            oDS_PH_PY105B.setValue "U_EXTAMT08", oRow, 0
            oDS_PH_PY105B.setValue "U_EXTAMT09", oRow, 0
            oDS_PH_PY105B.setValue "U_EXTAMT10", oRow, 0
            oMat1.LoadFromDataSource
        Else
            oDS_PH_PY105B.Offset = oRow - 1
            oDS_PH_PY105B.setValue "U_LineNum", oRow - 1, oRow
            oDS_PH_PY105B.setValue "U_JIGCOD", oRow - 1, ""
            oDS_PH_PY105B.setValue "U_HOBCOD", oRow - 1, ""
            oDS_PH_PY105B.setValue "U_HOBNAM", oRow - 1, ""
            oDS_PH_PY105B.setValue "U_STDAMT", oRow - 1, 0
            oDS_PH_PY105B.setValue "U_BNSAMT", oRow - 1, 0
            oDS_PH_PY105B.setValue "U_EXTAMT01", oRow - 1, 0
            oDS_PH_PY105B.setValue "U_EXTAMT02", oRow - 1, 0
            oDS_PH_PY105B.setValue "U_EXTAMT03", oRow - 1, 0
            oDS_PH_PY105B.setValue "U_EXTAMT04", oRow - 1, 0
            oDS_PH_PY105B.setValue "U_EXTAMT05", oRow - 1, 0
            oDS_PH_PY105B.setValue "U_EXTAMT06", oRow - 1, 0
            oDS_PH_PY105B.setValue "U_EXTAMT07", oRow - 1, 0
            oDS_PH_PY105B.setValue "U_EXTAMT08", oRow - 1, 0
            oDS_PH_PY105B.setValue "U_EXTAMT09", oRow - 1, 0
            oDS_PH_PY105B.setValue "U_EXTAMT10", oRow - 1, 0
            oMat1.LoadFromDataSource
        End If
    ElseIf oMat1.VisualRowCount = 0 Then
        oDS_PH_PY105B.Offset = oRow
        oDS_PH_PY105B.setValue "U_LineNum", oRow, oRow + 1
        oDS_PH_PY105B.setValue "U_JIGCOD", oRow, ""
        oDS_PH_PY105B.setValue "U_HOBCOD", oRow, ""
        oDS_PH_PY105B.setValue "U_HOBNAM", oRow, ""
        oDS_PH_PY105B.setValue "U_STDAMT", oRow, 0
        oDS_PH_PY105B.setValue "U_BNSAMT", oRow, 0
        oDS_PH_PY105B.setValue "U_EXTAMT01", oRow, 0
        oDS_PH_PY105B.setValue "U_EXTAMT02", oRow, 0
        oDS_PH_PY105B.setValue "U_EXTAMT03", oRow, 0
        oDS_PH_PY105B.setValue "U_EXTAMT04", oRow, 0
        oDS_PH_PY105B.setValue "U_EXTAMT05", oRow, 0
        oDS_PH_PY105B.setValue "U_EXTAMT06", oRow, 0
        oDS_PH_PY105B.setValue "U_EXTAMT07", oRow, 0
        oDS_PH_PY105B.setValue "U_EXTAMT08", oRow, 0
        oDS_PH_PY105B.setValue "U_EXTAMT09", oRow, 0
        oDS_PH_PY105B.setValue "U_EXTAMT10", oRow, 0
        oMat1.LoadFromDataSource
    End If
    
    Call oForm.Freeze(False)
    Exit Sub
PH_PY105_AddMatrixRow_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY105_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY105_FormClear()
On Error GoTo PH_PY105_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY105'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PH_PY105_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY105_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY105_DataValidCheck() As Boolean
On Error GoTo PH_PY105_DataValidCheck_Error
    PH_PY105_DataValidCheck = False
    Dim i       As Long
    Dim sQry    As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    '// 헤더 ---------------------------
    '//사업장
    If Trim(oDS_PH_PY105A.GetValue("U_CLTCOD", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "사업장은 필수입니다.", bmt_Short, True
        oForm.Items("CLTCOD").CLICK ct_Regular
        PH_PY105_DataValidCheck = False
        Exit Function
    End If
    
    '//적용시작월
    If Trim(oDS_PH_PY105A.GetValue("U_YM", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "적용시작월은 필수입니다.", bmt_Short, True
        oForm.Items("YM").CLICK ct_Regular
        PH_PY105_DataValidCheck = False
        Exit Function
    End If

    '// Code & Name 생성
    Call oDS_PH_PY105A.setValue("Code", 0, Trim(oDS_PH_PY105A.GetValue("U_CLTCOD", 0)) & Trim(oDS_PH_PY105A.GetValue("U_YM", 0)))
    Call oDS_PH_PY105A.setValue("NAME", 0, Trim(oDS_PH_PY105A.GetValue("U_CLTCOD", 0)) & Trim(oDS_PH_PY105A.GetValue("U_YM", 0)))
            
    If oForm.Mode = fm_ADD_MODE Then
        If MDC_SetMod.Get_ReData("Code", "Code", "[@PH_PY105A]", "'" & oDS_PH_PY105A.GetValue("Code", 0) & "'") <> "" Then
            Sbo_Application.SetStatusBarMessage "이미 존재하는 코드입니다.", bmt_Short, True
            Exit Function
        End If
    End If
    
    '// 라인 ---------------------------
    If oMat1.VisualRowCount > 1 Then
        For i = 1 To oMat1.VisualRowCount - 1
            '//호봉코드
            If oMat1.Columns("HOBCOD").Cells(i).Specific.VALUE = "" Then
                Sbo_Application.SetStatusBarMessage "호봉코드는 필수입니다.", bmt_Short, True
                oMat1.Columns("HOBCOD").Cells(i).CLICK ct_Regular
                PH_PY105_DataValidCheck = False
                Exit Function
            End If
            '//호봉명
            If oMat1.Columns("HOBNAM").Cells(i).Specific.VALUE = "" Then
                Sbo_Application.SetStatusBarMessage "내역 필수입니다.", bmt_Short, True
                oMat1.Columns("HOBNAM").Cells(i).CLICK ct_Regular
                PH_PY105_DataValidCheck = False
                Exit Function
            End If
        Next
    Else
        Sbo_Application.SetStatusBarMessage "라인 데이터가 없습니다.", bmt_Short, True
        PH_PY105_DataValidCheck = False
        Exit Function
    End If
    
    oMat1.FlushToDataSource
    
    '// Matrix 마지막 행 삭제(DB 저장시)
    If oDS_PH_PY105B.Size > 1 Then oDS_PH_PY105B.RemoveRecord (oDS_PH_PY105B.Size - 1)
    
    oMat1.LoadFromDataSource
    
    
    PH_PY105_DataValidCheck = True
    Set oRecordSet = Nothing
    Exit Function
    
    
    
    
PH_PY105_DataValidCheck_Error:
    Set oRecordSet = Nothing
    PH_PY105_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY105_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY105_MTX01()

    '//메트릭스에 데이터 로드
    
    Dim i       As Long
    Dim sQry    As String
    
    Dim Param01 As String
    Dim Param02 As String
    Dim Param03 As String
    Dim Param04 As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY105_MTX01_Error

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
        GoTo PH_PY105_MTX01_Exit
    End If
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet.RecordCount, False)
    
    For i = 0 To oRecordSet.RecordCount - 1
        If i <> 0 Then
            oDS_PH_PY105B.InsertRecord (i)
        End If
        oDS_PH_PY105B.Offset = i
        Call oDS_PH_PY105B.setValue("U_COL01", i, oRecordSet.Fields(0).VALUE)
        Call oDS_PH_PY105B.setValue("U_COL02", i, oRecordSet.Fields(1).VALUE)
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
PH_PY105_MTX01_Exit:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Exit Sub
PH_PY105_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY105_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY105_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY105_Validate_Error
    PH_PY105_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY105A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", bmt_Short, True
        PH_PY105_Validate = False
        GoTo PH_PY105_Validate_Exit
    End If
'
    If ValidateType = "수정" Then

    ElseIf ValidateType = "행삭제" Then

    ElseIf ValidateType = "취소" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY105_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY105_Validate_Error:
    PH_PY105_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY105_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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


Private Sub PH_PY105_Excel_Upload()
    
    Dim i           As Long
    Dim j           As Long
    Dim sPrice      As String
    Dim sFile       As String
    Dim OneRec      As String
    Dim sQry        As String
    
    Dim xl          As Excel.Application
    Dim xlwb        As Excel.Workbook
    Dim xlsh        As Excel.Worksheet
    
    Dim oEdit       As SAPbouiCOM.EditText
    Dim oForm       As SAPbouiCOM.Form
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo Err
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    Set oForm = Sbo_Application.Forms.ActiveForm
    
    
    '// 임시 테이블 생성
'    sQry = "EXEC MDC_MM_TEMP_CPD102 "
'
'    oRecordset.DoQuery sQry
    
    sFile = FileListBoxForm.OpenDialog(FileListBoxForm, "*.xls", "파일선택", "C:\")
    
    If sFile = "" Then
        Sbo_Application.StatusBar.SetText "파일을 선택해 주세요.", bmt_Short, smt_Error
        Exit Sub
    Else
        If Mid(Right(sFile, 4), 1, 3) = "xls" Or Mid(Right(sFile, 5), 1, 4) = "xlsx" Then
            Call oDS_PH_PY105A.setValue("U_Comments", 0, sFile)
        Else
            Sbo_Application.StatusBar.SetText "엑셀파일이 아닙니다.", bmt_Short, smt_Error
            Exit Sub
        End If
    End If
    
    '엑셀 Object 연결
    Set xl = CreateObject("excel.application")
    Set xlwb = xl.Workbooks.Open(sFile, , True)
    Set xlsh = xlwb.Worksheets("호봉등록")
    
    If xlsh.Cells(1, 1).VALUE <> "직급" Then
        Sbo_Application.StatusBar.SetText "A열 첫번째 행 타이틀은 직급", bmt_Short, smt_Error
        xlwb.Close
        Set xlwb = Nothing: Set xl = Nothing: Set xlsh = Nothing
        Exit Sub
    End If

    If xlsh.Cells(1, 2).VALUE <> "호봉코드" Then
        Sbo_Application.StatusBar.SetText "B열 두번째 행 타이틀은 호봉코드", bmt_Short, smt_Error
        xlwb.Close
        Set xlwb = Nothing: Set xl = Nothing: Set xlsh = Nothing
        Exit Sub
    End If

    If xlsh.Cells(1, 3).VALUE <> "호봉명" Then
        Sbo_Application.StatusBar.SetText "C열 세번째 행 타이틀은 호봉명", bmt_Short, smt_Error
        xlwb.Close
        Set xlwb = Nothing: Set xl = Nothing: Set xlsh = Nothing
        Exit Sub
    End If

    If xlsh.Cells(1, 4).VALUE <> "급여기본" Then
        Sbo_Application.StatusBar.SetText "D열 세번째 행 타이틀은 급여기본", bmt_Short, smt_Error
        xlwb.Close
        Set xlwb = Nothing: Set xl = Nothing: Set xlsh = Nothing
        Exit Sub
    End If

    If xlsh.Cells(1, 5).VALUE <> "상여기본" Then
        Sbo_Application.StatusBar.SetText "E열 세번째 행 타이틀은 상여기본", bmt_Short, smt_Error
        xlwb.Close
        Set xlwb = Nothing: Set xl = Nothing: Set xlsh = Nothing
        Exit Sub
    End If

    If xlsh.Cells(1, 6).VALUE <> "제수당01" Then
        Sbo_Application.StatusBar.SetText "F열 세번째 행 타이틀은 제수당01", bmt_Short, smt_Error
        xlwb.Close
        Set xlwb = Nothing: Set xl = Nothing: Set xlsh = Nothing
        Exit Sub
    End If
    
    If xlsh.Cells(1, 7).VALUE <> "제수당02" Then
        Sbo_Application.StatusBar.SetText "G열 세번째 행 타이틀은 제수당02", bmt_Short, smt_Error
        xlwb.Close
        Set xlwb = Nothing: Set xl = Nothing: Set xlsh = Nothing
        Exit Sub
    End If
    If xlsh.Cells(1, 8).VALUE <> "제수당03" Then
        Sbo_Application.StatusBar.SetText "H열 세번째 행 타이틀은 제수당03", bmt_Short, smt_Error
        xlwb.Close
        Set xlwb = Nothing: Set xl = Nothing: Set xlsh = Nothing
        Exit Sub
    End If
    If xlsh.Cells(1, 9).VALUE <> "제수당04" Then
        Sbo_Application.StatusBar.SetText "I열 세번째 행 타이틀은 제수당04", bmt_Short, smt_Error
        xlwb.Close
        Set xlwb = Nothing: Set xl = Nothing: Set xlsh = Nothing
        Exit Sub
    End If
    If xlsh.Cells(1, 10).VALUE <> "제수당05" Then
        Sbo_Application.StatusBar.SetText "J열 세번째 행 타이틀은 제수당05", bmt_Short, smt_Error
        xlwb.Close
        Set xlwb = Nothing: Set xl = Nothing: Set xlsh = Nothing
        Exit Sub
    End If
    If xlsh.Cells(1, 11).VALUE <> "제수당06" Then
        Sbo_Application.StatusBar.SetText "K열 세번째 행 타이틀은 제수당06", bmt_Short, smt_Error
        xlwb.Close
        Set xlwb = Nothing: Set xl = Nothing: Set xlsh = Nothing
        Exit Sub
    End If
    If xlsh.Cells(1, 12).VALUE <> "제수당07" Then
        Sbo_Application.StatusBar.SetText "L열 세번째 행 타이틀은 제수당07", bmt_Short, smt_Error
        xlwb.Close
        Set xlwb = Nothing: Set xl = Nothing: Set xlsh = Nothing
        Exit Sub
    End If
    If xlsh.Cells(1, 13).VALUE <> "제수당08" Then
        Sbo_Application.StatusBar.SetText "M열 세번째 행 타이틀은 제수당08", bmt_Short, smt_Error
        xlwb.Close
        Set xlwb = Nothing: Set xl = Nothing: Set xlsh = Nothing
        Exit Sub
    End If
    If xlsh.Cells(1, 14).VALUE <> "제수당09" Then
        Sbo_Application.StatusBar.SetText "N열 세번째 행 타이틀은 제수당09", bmt_Short, smt_Error
        xlwb.Close
        Set xlwb = Nothing: Set xl = Nothing: Set xlsh = Nothing
        Exit Sub
    End If
    If xlsh.Cells(1, 15).VALUE <> "제수당10" Then
        Sbo_Application.StatusBar.SetText "O열 세번째 행 타이틀은 제수당10", bmt_Short, smt_Error
        xlwb.Close
        Set xlwb = Nothing: Set xl = Nothing: Set xlsh = Nothing
        Exit Sub
    End If
    
    '//테이블 생성
    sQry = "EXEC PH_PY105_TEMP_CHK"
    oRecordSet.DoQuery sQry
    
    For i = 2 To xlsh.UsedRange.Rows.Count
        sQry = "EXEC PH_PY105 '" & xlsh.Cells(i, 1) & "','" & xlsh.Cells(i, 2) & "','" & xlsh.Cells(i, 3) & "','" & xlsh.Cells(i, 4) & "','" & xlsh.Cells(i, 5) & "','"
        sQry = sQry & xlsh.Cells(i, 6) & "','" & xlsh.Cells(i, 7) & "','" & xlsh.Cells(i, 8) & "','" & xlsh.Cells(i, 9) & "','" & xlsh.Cells(i, 10) & "','"
        sQry = sQry & xlsh.Cells(i, 11) & "','" & xlsh.Cells(i, 12) & "','" & xlsh.Cells(i, 13) & "','" & xlsh.Cells(i, 14) & "','" & xlsh.Cells(i, 15) & "'"
        
        oRecordSet.DoQuery sQry
    Next i
    
    oMat1.Clear
    oMat1.FlushToDataSource
    
    '// 임시데이터 데이타 검색
    sQry = "SELECT JIGCOD, HOBCOD, HOBNAM, STDAMT, BNSAMT, EXTAMT01, EXTAMT02, EXTAMT03, EXTAMT04, EXTAMT05, "
    sQry = sQry & " EXTAMT06, EXTAMT07, EXTAMT08, EXTAMT09, EXTAMT10 FROM PH_PY105_TEMP "
    oRecordSet.DoQuery sQry
    
    If oRecordSet.RecordCount > 0 Then
        For i = 0 To oRecordSet.RecordCount - 1
            oDS_PH_PY105B.InsertRecord (i)
            oDS_PH_PY105B.Offset = i
            oDS_PH_PY105B.setValue "U_LineNum", i, i + 1
            oDS_PH_PY105B.setValue "U_JIGCOD", i, oRecordSet.Fields(0).VALUE
            oDS_PH_PY105B.setValue "U_HOBCOD", i, oRecordSet.Fields(1).VALUE
            oDS_PH_PY105B.setValue "U_HOBNAM", i, oRecordSet.Fields(2).VALUE
            oDS_PH_PY105B.setValue "U_STDAMT", i, oRecordSet.Fields(3).VALUE
            oDS_PH_PY105B.setValue "U_BNSAMT", i, oRecordSet.Fields(4).VALUE
            oDS_PH_PY105B.setValue "U_EXTAMT01", i, oRecordSet.Fields(5).VALUE
            oDS_PH_PY105B.setValue "U_EXTAMT02", i, oRecordSet.Fields(6).VALUE
            oDS_PH_PY105B.setValue "U_EXTAMT03", i, oRecordSet.Fields(7).VALUE
            oDS_PH_PY105B.setValue "U_EXTAMT04", i, oRecordSet.Fields(8).VALUE
            oDS_PH_PY105B.setValue "U_EXTAMT05", i, oRecordSet.Fields(9).VALUE
            oDS_PH_PY105B.setValue "U_EXTAMT06", i, oRecordSet.Fields(10).VALUE
            oDS_PH_PY105B.setValue "U_EXTAMT07", i, oRecordSet.Fields(11).VALUE
            oDS_PH_PY105B.setValue "U_EXTAMT08", i, oRecordSet.Fields(12).VALUE
            oDS_PH_PY105B.setValue "U_EXTAMT09", i, oRecordSet.Fields(13).VALUE
            oDS_PH_PY105B.setValue "U_EXTAMT10", i, oRecordSet.Fields(14).VALUE
            oRecordSet.MoveNext
        Next i
        
    End If
    
    oMat1.LoadFromDataSource
    Call PH_PY105_AddMatrixRow
    
    Sbo_Application.StatusBar.SetText "엑셀을 불러왔습니다.", bmt_Short, smt_Warning


    '액셀개체 닫음
    xlwb.Close
    Set xlwb = Nothing: Set xl = Nothing: Set xlsh = Nothing
    
    '진행바 초기화
    Exit Sub

Err:
    xlwb.Close
    Set xlwb = Nothing: Set xl = Nothing: Set xlsh = Nothing
End Sub


Private Function PH_PY105_DataApply(CLTCOD As String, YM As String) As Boolean
    Dim sQry        As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    Dim Tablename As String
    Dim sTablename As String
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    On Error GoTo PH_PY105_DataApply_Error
    
    PH_PY105_DataApply = False
    
    oMat1.FlushToDataSource
    
    sTablename = "@PH_PY001A_" + YM '조회용
    
    
    sQry = " SELECT Count(*) FROM SYSOBJECTS WHERE Name = '" + sTablename + "'"
    oRecordSet.DoQuery sQry
    
    
    If oRecordSet.Fields(0).VALUE = 0 Then
        Tablename = "[@PH_PY001A_" + YM + "]" '//테이블 생성용
        sQry = "Exec PH_PY105_99 '" + Tablename + "'"
        oRecordSet.DoQuery sQry
    End If
    
    
    sQry = ""
    sQry = sQry & " Update T2 "
    sQry = sQry & " SET T2.U_STDAMT = T1.U_STDAMT, T2.U_BNSAMT = T1.U_BNSAMT, T2.U_HOBYMM = T0.U_YM"
    sQry = sQry & " FROM [@PH_PY105A] T0"
    sQry = sQry & " INNER JOIN [@PH_PY105B] T1 ON T0.Code = T1.Code"
    sQry = sQry & " INNER JOIN [@PH_PY001A] T2 ON T2.U_JIGCOD = T1.U_JIGCOD AND T2.U_HOBONG = T1.U_HOBCOD" ' AND T2.U_CLTCOD = T0.U_CLTCOD"
    sQry = sQry & " WHERE T0.U_YM = '" & YM & "'"
    sQry = sQry & " And T2.U_status <> '5' "
    sQry = sQry & " And Not Exists (Select * From [@PH_PY001A] T3 "
    sQry = sQry & " Where T2.Code = T3.Code"
    sQry = sQry & " And T3.U_status <> '5'"
    sQry = sQry & " And dbo.PH_PY_PAYPEAK_YEAR(T3.U_CLTCOD,'" & YM & "',T3.Code) > 0 )"  '호봉등록 년월에 임금피크제대상은 임금조정을 안함.
    'sQry = sQry & " AND T0.U_YM = '" & YM & "'"
    oRecordSet.DoQuery sQry
    
'    Sbo_Application.SetStatusBarMessage "해당 직급에 대헤 인사마스터에 금액이 적용 되었습니다." & Err.Description, bmt_Short, False
    Sbo_Application.StatusBar.SetText "해당 직급에 대헤 인사마스터에 금액이 적용 되었습니다.", bmt_Short, smt_Success
    Set oRecordSet = Nothing
    Exit Function
    
PH_PY105_DataApply_Error:
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "PH_PY105_DataApply_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

