VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY127"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY127.cls
'//  Module         : 급여관리 > 급여
'//  Desc           : 개인별 4대보험 보수월액 및 정산금액 등록(엑셀 Upload)
'//********************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form

Public oMat1 As SAPbouiCOM.Matrix

Private oDS_PH_PY127A As SAPbouiCOM.DBDataSource
Private oDS_PH_PY127B As SAPbouiCOM.DBDataSource

Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long

Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY127.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY127_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY127"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
    oForm.DataBrowser.BrowseBy = "Code"
        
    oForm.Freeze True
    Call PH_PY127_CreateItems
    Call PH_PY127_EnableMenus
    Call PH_PY127_SetDocument(oFromDocEntry01)
'    Call PH_PY127_FormResize
    
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

Private Function PH_PY127_CreateItems() As Boolean

    Dim sQry        As String
    Dim i           As Long
    Dim CLTCOD      As String
    
    Dim oCheck      As SAPbouiCOM.CheckBox
    Dim oEdit       As SAPbouiCOM.EditText
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns    As SAPbouiCOM.Columns
    Dim optBtn      As SAPbouiCOM.OptionBtn
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY127_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Set oDS_PH_PY127A = oForm.DataSources.DBDataSources("@PH_PY127A")
    Set oDS_PH_PY127B = oForm.DataSources.DBDataSources("@PH_PY127B")

    
    Set oMat1 = oForm.Items("Mat01").Specific       '//@PH_PY127B

    
    oMat1.SelectionMode = ms_NotSupported
    oMat1.AutoResizeColumns


    '//----------------------------------------------------------------------------------------------
    '// 기본사항
    '//----------------------------------------------------------------------------------------------
    
    '//구분(4대보험 선택)
    Set oCombo = oForm.Items("SIType").Specific
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P180' AND U_UseYN= 'Y'"
    Call SetReDataCombo(oForm, sQry, oCombo)


'
'    CLTCOD = MDC_SetMod.Get_ReData("Branch", "USER_CODE", "OUSR", "'" & oCompany.UserName & "'")
'    oCombo.Select CLTCOD, psk_ByValue
    
'    oForm.Items("CLTCOD").DisplayDesc = True
    
    '// 년
    Call oDS_PH_PY127A.setValue("U_YM", 0, Format(Now, "YYYYMM"))
    
    Call oForm.DataSources.UserDataSources.Add("Comments", dt_SHORT_TEXT, 100)
    oForm.Items("Comments").Specific.DataBind.SetBound True, "", "Comments"
    
        
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
    
PH_PY127_CreateItems_Error:
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY127_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY127_EnableMenus()
    
    On Error GoTo PH_PY127_EnableMenus_Error

    Call oForm.EnableMenu("1283", True)     '//제거
    Call oForm.EnableMenu("1284", False)    '//취소
    Call oForm.EnableMenu("1293", False)     '//행삭제
    
    Exit Sub
    
PH_PY127_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY127_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY127_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY127_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY127_FormItemEnabled
'        Call PH_PY127_AddMatrixRow
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY127_FormItemEnabled
        oForm.Items("Code").Specific.VALUE = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY127_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY127_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY127_FormItemEnabled()
    Dim oCombo      As SAPbouiCOM.ComboBox
    
    On Error GoTo PH_PY127_FormItemEnabled_Error
    
    
    Call oForm.Freeze(True)
    If (oForm.Mode = fm_ADD_MODE) Then
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        
        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", False)    '//문서추가
        Call oForm.EnableMenu("1293", False)    '//행삭제
        
        oForm.Items("Code").Enabled = False
        oForm.Items("YM").Enabled = True
        oForm.Items("CLTCOD").Enabled = True
        oForm.Items("SIType").Enabled = True
        oForm.Items("Comments").Specific.VALUE = ""
    ElseIf (oForm.Mode = fm_FIND_MODE) Then
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        
        Call oForm.EnableMenu("1281", False)    '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가
        Call oForm.EnableMenu("1293", False)    '//행삭제
        
        oForm.Items("Code").Enabled = True
        oForm.Items("YM").Enabled = True
        oForm.Items("CLTCOD").Enabled = True
        oForm.Items("SIType").Enabled = True
        oForm.Items("Comments").Specific.VALUE = ""
        
    ElseIf (oForm.Mode = fm_OK_MODE) Then
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD", False)
        
        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가
        Call oForm.EnableMenu("1293", False)    '//행삭제
        
        oForm.Items("Code").Enabled = False
        oForm.Items("YM").Enabled = False
        oForm.Items("CLTCOD").Enabled = False
        oForm.Items("SIType").Enabled = False
        oForm.Items("Comments").Specific.VALUE = ""
    End If
    Call oForm.Freeze(False)
    Exit Sub
    
PH_PY127_FormItemEnabled_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY127_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Public Sub Raise_FormItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
    Dim sQry        As String
    Dim i           As Long
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Dim Code As String
    Dim YM As String
    
    On Error GoTo Raise_FormItemEvent_Error
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Select Case pval.EventType
        Case et_ITEM_PRESSED: '//1
            
            If pval.BeforeAction = True Then
                If pval.ItemUID = "1" Then
                    If oForm.Mode = fm_ADD_MODE Then
                        If PH_PY127_DataValidCheck("Y") = False Then
                            BubbleEvent = False
                        End If
                        
                        '//해야할일 작업
                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
                        '//해야할일 작업
                        
                    ElseIf oForm.Mode = fm_OK_MODE Then
                        
                    End If
                End If
                
            ElseIf pval.BeforeAction = False Then
                If pval.ItemUID = "1" Then
                    If oForm.Mode = fm_ADD_MODE Then
                        If pval.ActionSuccess = True Then
                            
'                            sQry = "Delete [@PH_PY127B] From [@PH_PY127A] a Where [@PH_PY127B].Code = a.Code "
'                            sQry = sQry + " and Isnull([@PH_PY127B].U_CLTCOD,'') = ''"
'                            oRecordSet.DoQuery sQry
                            
                            Call PH_PY127_FormItemEnabled
                            
                            
    '                        Call PH_PY127_AddMatrixRow
                        End If
                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY127_FormItemEnabled
    '                        Call PH_PY127_AddMatrixRow
                        End If
                    ElseIf oForm.Mode = fm_OK_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY127_FormItemEnabled
                        End If
                    End If
                End If
                If pval.ItemUID = "Btn_UPLOAD" Then
                    Call PH_PY127_Excel_Upload
                End If
                If pval.ItemUID = "Btn_UPDATE" Then
                    Call PH_PY127_PY001A_UPDATE
                End If
                If pval.ItemUID = "Btn_CANCEL" Then
                    Call PH_PY127_PY001A_CANCEL
                End If

            End If
     
'----------------------------------------------------------
        Case et_KEY_DOWN: '//2

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
       
            ElseIf pval.BeforeAction = False Then
                If pval.ItemChanged = True Then

                End If
            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_MATRIX_LOAD: '//11
            If pval.BeforeAction = True Then
            ElseIf pval.BeforeAction = False Then
                oMat1.LoadFromDataSource
                Call PH_PY127_FormItemEnabled
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
                Set oDS_PH_PY127A = Nothing
                Set oDS_PH_PY127B = Nothing
                
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
'
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
'                    Call MDC_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY127A", "Code")
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
    Call oForm.Freeze(True)
    
    If (pval.BeforeAction = True) Then
        Select Case pval.MenuUID
            Case "1283":
                '보수월액 적용을 하면 삭제를 못하도록, 적용취소후 삭제 가능
                
                If Val(oMat1.Columns("BeforAmt").Cells(1).Specific.VALUE) > 0 Then
                    Sbo_Application.StatusBar.SetText "보수월액 갱신처리한 자료입니다. 삭제할 수 없습니다.", bmt_Short, smt_Error
                    Call oForm.Freeze(False)
                    BubbleEvent = False
                    Exit Sub
                End If
                
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
                'Call AuthorityCheck(oForm, "CLTCOD", "@PH_PY127A", "Code")      '//접속자 권한에 따른 사업장 보기
        End Select
    ElseIf (pval.BeforeAction = False) Then
        Select Case pval.MenuUID
            Case "1283":
                oForm.Mode = fm_ADD_MODE
                Call PH_PY127_FormItemEnabled
'                Call PH_PY127_AddMatrixRow
            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//문서찾기
                Call PH_PY127_FormItemEnabled
'                Call PH_PY127_AddMatrixRow
                oForm.Items("Code").CLICK ct_Regular
            Case "1282":    '//문서추가
                Call PH_PY127_FormItemEnabled
'                Call PH_PY127_AddMatrixRow
            Case "1288", "1289", "1290", "1291":
                Call PH_PY127_FormItemEnabled
            Case "1293" '// 행삭제
                
'                Call PH_PY127_AddMatrixRow
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

Sub PH_PY127_AddMatrixRow()
    Dim oRow    As Long
    
    On Error GoTo PH_PY127_AddMatrixRow_Error
    
    Call oForm.Freeze(True)
    
    '//[Mat1]
    oMat1.FlushToDataSource
    oRow = oMat1.VisualRowCount
    
    If oMat1.VisualRowCount > 0 Then
        If Trim(oDS_PH_PY127B.GetValue("U_FullName", oRow - 1)) <> "" Then
            If oDS_PH_PY127B.Size <= oMat1.VisualRowCount Then
                oDS_PH_PY127B.InsertRecord (oRow)
            End If
            oDS_PH_PY127B.Offset = oRow
            oDS_PH_PY127B.setValue "U_LineNum", oRow, oRow + 1
            oDS_PH_PY127B.setValue "U_FullName", oRow, ""
            oDS_PH_PY127B.setValue "U_govID", oRow, ""
            oDS_PH_PY127B.setValue "U_Amt", oRow, 0
            oDS_PH_PY127B.setValue "AjustAmt", oRow, 0
            oDS_PH_PY127B.setValue "U_MSTCOD", oRow, ""
            oDS_PH_PY127B.setValue "BeforAmt", oRow, 0
            oMat1.LoadFromDataSource
        Else
            oDS_PH_PY127B.Offset = oRow - 1
            oDS_PH_PY127B.setValue "U_LineNum", oRow, oRow + 1
            oDS_PH_PY127B.setValue "U_FullName", oRow, ""
            oDS_PH_PY127B.setValue "U_govID", oRow, ""
            oDS_PH_PY127B.setValue "U_Amt", oRow, 0
            oDS_PH_PY127B.setValue "AjustAmt", oRow, 0
            oDS_PH_PY127B.setValue "U_MSTCOD", oRow, ""
            oDS_PH_PY127B.setValue "BeforAmt", oRow, 0
            oMat1.LoadFromDataSource
        End If
    ElseIf oMat1.VisualRowCount = 0 Then
        oDS_PH_PY127B.Offset = oRow
            oDS_PH_PY127B.setValue "U_LineNum", oRow, oRow + 1
            oDS_PH_PY127B.setValue "U_FullName", oRow, ""
            oDS_PH_PY127B.setValue "U_govID", oRow, ""
            oDS_PH_PY127B.setValue "U_Amt", oRow, 0
            oDS_PH_PY127B.setValue "AjustAmt", oRow, 0
            oDS_PH_PY127B.setValue "U_MSTCOD", oRow, ""
            oDS_PH_PY127B.setValue "BeforAmt", oRow, 0
        oMat1.LoadFromDataSource
    End If
    
    Call oForm.Freeze(False)
    Exit Sub
PH_PY127_AddMatrixRow_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY127_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY127_FormClear()
On Error GoTo PH_PY127_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY127'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PH_PY127_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY127_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY127_DataValidCheck(ChkYN As String) As Boolean
On Error GoTo PH_PY127_DataValidCheck_Error
    PH_PY127_DataValidCheck = False
    Dim i       As Long
    Dim sQry    As String
    Dim tCode   As String
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    If ChkYN = "Y" Then
        If Trim(oDS_PH_PY127A.GetValue("U_CLTCOD", 0)) = "" Then
            Sbo_Application.SetStatusBarMessage "사업장은 필수입니다.", bmt_Short, True
            oForm.Items("CLTCOD").CLICK ct_Regular
            PH_PY127_DataValidCheck = False
            Exit Function
        End If
        
        If Trim(oDS_PH_PY127A.GetValue("U_YM", 0)) = "" Then
            Sbo_Application.SetStatusBarMessage "년월은 필수입니다.", bmt_Short, True
            oForm.Items("YM").CLICK ct_Regular
            PH_PY127_DataValidCheck = False
            Exit Function
        End If
        If Trim(oDS_PH_PY127A.GetValue("U_SIType", 0)) = "" Then
            Sbo_Application.SetStatusBarMessage "구분은 필수입니다.", bmt_Short, True
            oForm.Items("SIType").CLICK ct_Regular
            PH_PY127_DataValidCheck = False
            Exit Function
        End If
    End If
    
    
    '// 코드,이름 저장
    tCode = Trim(oDS_PH_PY127A.GetValue("U_YM", 0)) + Trim(oDS_PH_PY127A.GetValue("U_SIType", 0)) + Trim(oDS_PH_PY127A.GetValue("U_CLTCOD", 0))
    Call oDS_PH_PY127A.setValue("Code", 0, tCode)
    Call oDS_PH_PY127A.setValue("Name", 0, tCode)
    
    If oForm.Mode = fm_ADD_MODE Then
        '// 데이터 중복 체크
        sQry = "SELECT Code FROM [@PH_PY127A] WHERE Code = '" & tCode & "'"
        oRecordSet.DoQuery sQry
        
        If oRecordSet.RecordCount > 0 Then
            Sbo_Application.SetStatusBarMessage "이미 데이터가 존재합니다.", bmt_Short, True
            PH_PY127_DataValidCheck = False
            Exit Function
        End If
    End If
    
    If ChkYN = "Y" Then
        If oMat1.VisualRowCount = 0 Then
            Sbo_Application.SetStatusBarMessage "데이터가 없습니다. 확인바랍니다", bmt_Short, True
            PH_PY127_DataValidCheck = False
            Exit Function
        End If
    End If
    
    PH_PY127_DataValidCheck = True
    
    Exit Function
    
    Set oRecordSet = Nothing
    
    
PH_PY127_DataValidCheck_Error:
    Set oRecordSet = Nothing
    PH_PY127_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY127_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function



Function PH_PY127_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY127_Validate_Error
    PH_PY127_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY127A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", bmt_Short, True
        PH_PY127_Validate = False
        GoTo PH_PY127_Validate_Exit
    End If
'
    If ValidateType = "수정" Then

    ElseIf ValidateType = "행삭제" Then

    ElseIf ValidateType = "취소" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY127_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY127_Validate_Error:
    PH_PY127_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY127_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY127_Excel_Upload()
    
    Dim i           As Long
    Dim j           As Long
    Dim sPrice      As String
    Dim sFile       As String
    Dim OneRec      As String
    Dim sQry        As String
    Dim FullName    As String
    Dim GovID       As String
    Dim CLTCOD      As String
    
    Dim xl          As Excel.Application
    Dim xlwb        As Excel.Workbook
    Dim xlsh        As Excel.Worksheet
    
    Dim oEdit       As SAPbouiCOM.EditText
    Dim oForm       As SAPbouiCOM.Form
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo Err
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    Set oForm = Sbo_Application.Forms.ActiveForm
    
    
    CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE)
    
    
    sFile = FileListBoxForm.OpenDialog(FileListBoxForm, "*.xls", "파일선택", "C:\")
    
    If sFile = "" Then
        Sbo_Application.StatusBar.SetText "파일을 선택해 주세요.", bmt_Short, smt_Error
        Exit Sub
    Else
        If Mid(Right(sFile, 4), 1, 3) = "xls" Or Mid(Right(sFile, 5), 1, 4) = "xlsx" Then
            oForm.Items("Comments").Specific.VALUE = sFile

        Else
            Sbo_Application.StatusBar.SetText "엑셀파일이 아닙니다.", bmt_Short, smt_Error
            Exit Sub
        End If
    End If
    
    '엑셀 Object 연결
    Set xl = CreateObject("excel.application")
    Set xlwb = xl.Workbooks.Open(sFile, , True)
    Set xlsh = xlwb.Worksheets("Sheet1")
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    'Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet.RecordCount, False)
'    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", 0, False)
    
    Call oForm.Freeze(True)
    
    oMat1.Clear
    oMat1.FlushToDataSource
    oMat1.LoadFromDataSource
    For i = 2 To xlsh.UsedRange.Rows.Count
        
        If i - 2 <> 0 Then
            oDS_PH_PY127B.InsertRecord (i - 2)
        End If
        oDS_PH_PY127B.Offset = i - 2
        Call oDS_PH_PY127B.setValue("U_LineNum", i - 2, i - 1)
        Call oDS_PH_PY127B.setValue("U_FullName", i - 2, xlsh.Cells(i, 1))       '성명
        Call oDS_PH_PY127B.setValue("U_govID", i - 2, xlsh.Cells(i, 2))          '주민번호
        Call oDS_PH_PY127B.setValue("U_Amt", i - 2, xlsh.Cells(i, 3))            '보수월액
        Call oDS_PH_PY127B.setValue("U_AjustAmt", i - 2, Val(xlsh.Cells(i, 4)))  '정산금액
        
'        ProgressBar01.VALUE = ProgressBar01.VALUE + 1
'        ProgressBar01.Text = ProgressBar01.VALUE & "/" & xlsh.UsedRange.Rows.Count - 1 & "건 조회중...!"
    Next i
    oMat1.LoadFromDataSource
    oMat1.AutoResizeColumns
    oForm.Update

    Call PH_PY127_AddMatrixRow
    
    
    
    Sbo_Application.StatusBar.SetText "엑셀을 불러왔습니다.", bmt_Short, smt_Warning
    
    For i = 1 To oMat1.VisualRowCount
        FullName = oMat1.Columns("FullName").Cells(i).Specific.VALUE
        GovID = Left(oMat1.Columns("govID").Cells(i).Specific.VALUE, 6) + Right(oMat1.Columns("govID").Cells(i).Specific.VALUE, 7)


        sQry = "Select Code From [@PH_PY001A] wHERE U_status <> '5' and U_CLTCOD = '" & CLTCOD & "' and U_FullName = '" & FullName & "' And U_govID = '" & GovID & "'"
        oRecordSet.DoQuery sQry
        If oRecordSet.RecordCount > 0 Then
            oMat1.Columns("MSTCOD").Cells(i).Specific.VALUE = oRecordSet.Fields(0).VALUE
            oMat1.Columns("govID").Cells(i).Specific.VALUE = Left(GovID, 6) + "-" + Mid(GovID, 7, 1) + "******"
        Else
            oMat1.Columns("govID").Cells(i).Specific.VALUE = Left(GovID, 6) + "-" + Mid(GovID, 7, 1) + "******"
        End If
    Next i
    
    Call oForm.Freeze(False)
    
    '액셀개체 닫음
    xlwb.Close
    Set xlwb = Nothing: Set xl = Nothing: Set xlsh = Nothing
'    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    '진행바 초기화
    Exit Sub

Err:
    xlwb.Close
    Set xlwb = Nothing: Set xl = Nothing: Set xlsh = Nothing
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
End Sub

Private Sub PH_PY127_PY001A_UPDATE()
    
    Dim sQry        As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    Dim Code        As String
    Dim CLTCOD      As String
    Dim YM          As String
    Dim SIType      As String
    
    On Error GoTo Err
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If oForm.Mode = fm_OK_MODE Then
        
        CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE)
        YM = oForm.Items("YM").Specific.VALUE
        Code = oForm.Items("Code").Specific.VALUE
        SIType = oForm.Items("SIType").Specific.VALUE
        
        If oMat1.VisualRowCount <= 1 Then
            Sbo_Application.StatusBar.SetText "처리할 라인데이타가 없습니다. 확인바랍니다.", bmt_Short, smt_Error
            Exit Sub
        End If
        
        If Val(oMat1.Columns("BeforAmt").Cells(1).Specific.VALUE) > 0 Then
            Sbo_Application.StatusBar.SetText "보수월액 갱신처리한 자료입니다. 확인바랍니다.", bmt_Short, smt_Error
            Exit Sub
        End If
        
        sQry = "Select Cnt = Count(*) From [@PH_PY127A] Where U_YM > '" & YM & "' and U_SIType = '" & SIType & "' and U_CLTCOD = '" & CLTCOD & "'"
        oRecordSet.DoQuery sQry
        
        If oRecordSet.Fields(0).VALUE > 0 Then
            Sbo_Application.StatusBar.SetText "보수월액 처리 이후 작업자료가 있습니다. 이전 자료로는 작업 할 수 없습니다.", bmt_Short, smt_Error
            Exit Sub
        End If
        
                
        sQry = "Exec PH_PY127_01 '" & CLTCOD & "', '" & Code & "'"
        oRecordSet.DoQuery sQry
        
        Call Sbo_Application.ActivateMenuItem("1281")
        oForm.Items("Code").Specific.VALUE = Code
        oForm.Items("1").CLICK ct_Regular
        
    Else
        Sbo_Application.StatusBar.SetText "해당모드에서는 작업할 수 없습니다. 작업할 문서를 불러와서 작업 바랍니다.", bmt_Short, smt_Error
    End If
    
    
    
    Exit Sub

Err:
    
End Sub

Private Sub PH_PY127_PY001A_CANCEL()
    
    Dim sQry        As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    Dim Code        As String
    Dim CLTCOD      As String
    
    
    On Error GoTo Err
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If oForm.Mode = fm_OK_MODE Then
        If oMat1.VisualRowCount <= 1 Then
            Sbo_Application.StatusBar.SetText "처리할 라인데이타가 없습니다. 확인바랍니다.", bmt_Short, smt_Error
            Exit Sub
        End If
        
        If Val(oMat1.Columns("BeforAmt").Cells(1).Specific.VALUE) = 0 Then
            Sbo_Application.StatusBar.SetText "보수월액 적용전 자료입니다. 작업취소를 할 수 없습니다.", bmt_Short, smt_Error
            Exit Sub
        End If
        
        
        
        CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE)
        Code = oForm.Items("Code").Specific.VALUE
        
        sQry = "Exec PH_PY127_02 '" & CLTCOD & "', '" & Code & "'"
        oRecordSet.DoQuery sQry
        
        Call Sbo_Application.ActivateMenuItem("1281")
        oForm.Items("Code").Specific.VALUE = Code
        oForm.Items("1").CLICK ct_Regular
        
    Else
        Sbo_Application.StatusBar.SetText "해당모드에서는 작업할 수 없습니다. 작업할 문서를 불러와서 작업 바랍니다.", bmt_Short, smt_Error
    End If
    
    
    
    Exit Sub

Err:
    
End Sub

''

