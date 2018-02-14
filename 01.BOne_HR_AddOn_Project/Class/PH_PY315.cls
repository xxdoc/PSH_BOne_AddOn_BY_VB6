VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY315"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY315.cls
'//  Module         : 인사관리 > 기타관리 > 대부금관리
'//  Desc           : 개인별 대부금 잔액현황
'//********************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form

Public oMat1 As SAPbouiCOM.Matrix
Public oMat2 As SAPbouiCOM.Matrix

Private oDS_PH_PY315A As SAPbouiCOM.DBDataSource
Private oDS_PH_PY315B As SAPbouiCOM.DBDataSource
Private oDS_PH_PY315C As SAPbouiCOM.DBDataSource

Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long

Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY315.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY315_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY315"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
'    oForm.DataBrowser.BrowseBy = "DocEntry"
        
    oForm.Freeze True
    Call PH_PY315_CreateItems
    Call PH_PY315_EnableMenus
    Call PH_PY315_SetDocument(oFromDocEntry01)
    Call PH_PY315_FormResize
    
    oForm.Update
    oForm.Freeze False
    
'    oForm.Items("Folder01").Specific.Select '폼이 로드 될 때 Folder01이 선택됨
    
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

Private Function PH_PY315_CreateItems() As Boolean

    Dim sQry        As String
    Dim i           As Long
    
    Dim oEdit       As SAPbouiCOM.EditText
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns    As SAPbouiCOM.Columns
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY315_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
'    Set oDS_PH_PY315A = oForm.DataSources.DBDataSources("@PH_PY315A")
    Set oDS_PH_PY315B = oForm.DataSources.DBDataSources("@PS_USERDS01")
'    Set oDS_PH_PY315C = oForm.DataSources.DBDataSources("@PS_USERDS02")
    
    Set oMat1 = oForm.Items("Mat01").Specific
'    Set oMat2 = oForm.Items("Mat02").Specific

    oMat1.SelectionMode = ms_Auto
    oMat1.AutoResizeColumns
    
'    oMat2.SelectionMode = ms_Auto
'    oMat2.AutoResizeColumns

    '----------조회 조건----------
    '사업장_S
    Call oForm.DataSources.UserDataSources.Add("CLTCOD", dt_SHORT_TEXT, 5)
    Call oForm.Items("CLTCOD").Specific.DataBind.SetBound(True, "", "CLTCOD")
    '사업장_E
    
    '상환일자_S
    Call oForm.DataSources.UserDataSources.Add("RpmtDate", dt_DATE)
    Call oForm.Items("RpmtDate").Specific.DataBind.SetBound(True, "", "RpmtDate")
    '상환일자_E
    
    '사번_S
    Call oForm.DataSources.UserDataSources.Add("CntcCode", dt_SHORT_TEXT, 20)
    Call oForm.Items("CntcCode").Specific.DataBind.SetBound(True, "", "CntcCode")
    '사번_E
    
    '성명_S
    Call oForm.DataSources.UserDataSources.Add("CntcName", dt_SHORT_TEXT, 50)
    Call oForm.Items("CntcName").Specific.DataBind.SetBound(True, "", "CntcName")
    '성명_E
    '----------조회 조건----------
    
    '//----------------------------------------------------------------------------------------------
    '// 기본사항
    '//----------------------------------------------------------------------------------------------
    
    '사업장
    Set oCombo = oForm.Items("CLTCOD").Specific
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
    Call SetReDataCombo(oForm, sQry, oCombo)
    oForm.Items("CLTCOD").DisplayDesc = True
    
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
    
PH_PY315_CreateItems_Error:
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY315_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY315_EnableMenus()
    
    On Error GoTo PH_PY315_EnableMenus_Error
    
'    Call oForm.EnableMenu("1283", False)         '// 삭제
'    Call oForm.EnableMenu("1287", True)          '// 복제
''    Call oForm.EnableMenu("1286", True)         '// 닫기
'    Call oForm.EnableMenu("1284", True)         '// 취소
'    Call oForm.EnableMenu("1293", True)         '// 행삭제
    
    Exit Sub
    
PH_PY315_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY315_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY315_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY315_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY315_FormItemEnabled
'        Call PH_PY315_AddMatrixRow1
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY315_FormItemEnabled
        oForm.Items("DocEntry").Specific.VALUE = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY315_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY315_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY315_FormItemEnabled()
    On Error GoTo PH_PY315_FormItemEnabled_Error
    
    Dim oCombo As SAPbouiCOM.ComboBox
    Dim CLTCOD As String
        
    Call oForm.Freeze(True)
    If (oForm.Mode = fm_ADD_MODE) Then

        '폼 DocEntry 세팅
'        Call PH_PY315_FormClear
        
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        
        '일자 세팅
        oForm.Items("RpmtDate").Specific.VALUE = Format(Date, "YYYYMMDD")

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
    
PH_PY315_FormItemEnabled_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY315_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub Raise_FormItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
    On Error GoTo Raise_FormItemEvent_Error
    
    Dim sQry As String
    Dim i As Long
    Dim oCombo As SAPbouiCOM.ComboBox
    Dim oRecordSet As SAPbobsCOM.Recordset
    
'    Dim loopCount As Integer 'For Loop 용 (VALIDATE Event에서 사용)
'    Dim GovID1 As String '주민등록번호 앞자리(VALIDATE Event에서 사용)
'    Dim GovID2 As String '주민등록번호 뒷자리(VALIDATE Event에서 사용)
'    Dim GovID As String '주민등록번호 전체(VALIDATE Event에서 사용)
'    Dim Sex As String '성별(VALIDATE Event에서 사용)
'    Dim PayCnt As Integer '지급횟수(VALIDATE Event에서 사용)
'    Dim FeeTot As Double '입학금계(VALIDATE Event에서 사용)
'    Dim TuiTot As Double '등록금계(VALIDATE Event에서 사용)
'    Dim Total As Double '총계(VALIDATE Event에서 사용)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Select Case pval.EventType
        Case et_ITEM_PRESSED: '//1
            
            If pval.BeforeAction = True Then
'                If pval.ItemUID = "1" Then
'                    If oForm.Mode = fm_ADD_MODE Then
'                        If PH_PY315_DataValidCheck = False Then
'                            BubbleEvent = False
'                        End If
'
'                        '//해야할일 작업
'                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
'                        If PH_PY315_DataValidCheck = False Then
'                            BubbleEvent = False
'                        End If
'                        '//해야할일 작업
'
'                    ElseIf oForm.Mode = fm_OK_MODE Then
'                    End If
'                End If
                If pval.ItemUID = "BtnSearch" Then
                
                    If oForm.Mode = fm_ADD_MODE Then
                        If PH_PY315_DataValidCheck = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
'
'                        '//해야할일 작업
                        Call PH_PY315_MTX01
'                        Call PH_PY315_MTX02
                        
                    End If
               
                End If
                
                If pval.ItemUID = "BtnPrint1" Then
                
                    If oForm.Mode = fm_ADD_MODE Then
                        If PH_PY315_DataValidCheck = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
'
'                        '//해야할일 작업
                        Call PH_PY315_Print_Report01
                        
                    End If
               
                End If
                
            ElseIf pval.BeforeAction = False Then
'                If pval.ItemUID = "1" Then
'                    If oForm.Mode = fm_ADD_MODE Then
'                        If pval.ActionSuccess = True Then
'                            Call PH_PY315_FormItemEnabled
'                            Call PH_PY315_AddMatrixRow1
'                        End If
'                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
'                        If pval.ActionSuccess = True Then
'                            Call PH_PY315_FormItemEnabled
'                            Call PH_PY315_AddMatrixRow1
'                        End If
'                    ElseIf oForm.Mode = fm_OK_MODE Then
'                        If pval.ActionSuccess = True Then
'                            Call PH_PY315_FormItemEnabled
'                        End If
'                    End If
'                End If

'                '폴더를 사용할 때는 필수 소스_S
'                'Folder01이 선택되었을 때
'                If pval.ItemUID = "Folder01" Then
'
'                    oForm.PaneLevel = 1
'
'                End If
'
'                'Folder02가 선택되었을 때
'                If pval.ItemUID = "Folder02" Then
'
'                    oForm.PaneLevel = 2
'
'                End If
'                '폴더를 사용할 때는 필수 소스_E
                
            End If
     
'----------------------------------------------------------
        Case et_KEY_DOWN: '//2

            If pval.BeforeAction = True Then
            
                If pval.ItemUID = "Mat01" Then
        
                    If pval.ColUID = "Name" And pval.CharPressed = "9" Then
                    
                        If oMat1.Columns.Item("Name").Cells(pval.Row).Specific.VALUE = "" Then
                            Call Sbo_Application.ActivateMenuItem("7425")
                            BubbleEvent = False
                        End If
                    
                    End If
                
                ElseIf pval.ItemUID = "CntcCode" And pval.CharPressed = "9" Then
                
                    If oForm.Items("CntcCode").Specific.VALUE = "" Then
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
'                    Call PH_PY315_AddMatrixRow1
                    Call oMat1.AutoResizeColumns
                End If
            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_CLICK: '//6
            If pval.BeforeAction = True Then
'                Select Case pval.ItemUID
'                Case "Mat01"
'                    If pval.Row > 0 Then
'                        Call oMat1.SelectRow(pval.Row, True, False)
'                    End If
'                End Select
                
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
                            'Call oDS_PH_PY315A.setValue("U_CntcName", 0, MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & oForm.Items(pval.ItemUID).Specific.Value & "'"))
                        
                            oForm.Items("CntcName").Specific.VALUE = MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & oForm.Items(pval.ItemUID).Specific.VALUE & "'")
                        
                    End Select

                End If
                
            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_MATRIX_LOAD: '//11
            If pval.BeforeAction = True Then
            ElseIf pval.BeforeAction = False Then
                oMat1.LoadFromDataSource

                Call PH_PY315_FormItemEnabled
'                Call PH_PY315_AddMatrixRow1
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
                Set oDS_PH_PY315A = Nothing
                Set oDS_PH_PY315B = Nothing
                Set oDS_PH_PY315C = Nothing
                
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

                Call PH_PY315_FormResize
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
'                    Call MDC_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY315A", "Code")
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
            
            Case "7169": '엑셀 내보내기
            
                '엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
                Call PH_PY315_AddMatrixRow1
            
        End Select
    ElseIf (pval.BeforeAction = False) Then
        Select Case pval.MenuUID
            Case "1283":
                oForm.Mode = fm_ADD_MODE
                Call PH_PY315_FormItemEnabled
                Call PH_PY315_AddMatrixRow1
            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//문서찾기
                Call PH_PY315_FormItemEnabled
                Call PH_PY315_AddMatrixRow1
                oForm.Items("DocEntry").CLICK ct_Regular
            Case "1282":    '//문서추가
                Call PH_PY315_FormItemEnabled
                Call PH_PY315_AddMatrixRow1
            Case "1288", "1289", "1290", "1291":
                Call PH_PY315_FormItemEnabled
            Case "1293" '// 행삭제

                 If oMat1.RowCount <> oMat1.VisualRowCount Then
                    oMat1.FlushToDataSource
    
                    While (i <= oDS_PH_PY315B.Size - 1)
                        If oDS_PH_PY315B.GetValue("U_LineNum", i) = "" Then
                            oDS_PH_PY315B.RemoveRecord (i)
                            i = 0
                        Else
                            i = i + 1
                        End If
                    Wend
                    
                    For i = 0 To oDS_PH_PY315B.Size
                        Call oDS_PH_PY315B.setValue("U_LineNum", i, i + 1)
                    Next i
                    
                    oMat1.LoadFromDataSource
                End If
                Call PH_PY315_AddMatrixRow1
                
            Case "7169": '엑셀 내보내기
            
                '엑셀 내보내기 이후 처리
                Call oForm.Freeze(True)
                Call oDS_PH_PY315B.RemoveRecord(oDS_PH_PY315B.Size - 1)
                Call oMat1.LoadFromDataSource
                Call oForm.Freeze(False)
                
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

Sub PH_PY315_AddMatrixRow1()
    Dim oRow    As Long
    
    On Error GoTo PH_PY315_AddMatrixRow1_Error
    
    Call oForm.Freeze(True)
    
    '//[Mat1]
    oMat1.FlushToDataSource
    oRow = oMat1.VisualRowCount
    
    If oMat1.VisualRowCount > 0 Then
        If Trim(oDS_PH_PY315B.GetValue("U_LineNum", oRow - 1)) <> "" Then
            If oDS_PH_PY315B.Size <= oMat1.VisualRowCount Then
                oDS_PH_PY315B.InsertRecord (oRow)
            End If
            oDS_PH_PY315B.Offset = oRow
            oDS_PH_PY315B.setValue "U_LineNum", oRow, oRow + 1
            oDS_PH_PY315B.setValue "U_ColReg01", oRow, "" '부서
            oDS_PH_PY315B.setValue "U_ColReg02", oRow, "" '직급
            oDS_PH_PY315B.setValue "U_ColReg03", oRow, "" '사번
            oDS_PH_PY315B.setValue "U_ColReg04", oRow, "" '성명
            oDS_PH_PY315B.setValue "U_ColDt01", oRow, "" '대출일자
            oDS_PH_PY315B.setValue "U_ColSum01", oRow, 0 '신규대출액
            oDS_PH_PY315B.setValue "U_ColQty01", oRow, 0 '금리
            oDS_PH_PY315B.setValue "U_ColSum02", oRow, 0 '대출총액
            oDS_PH_PY315B.setValue "U_ColSum03", oRow, 0 '전월대출잔액
            oDS_PH_PY315B.setValue "U_ColSum04", oRow, 0 '당월상환원금
            oDS_PH_PY315B.setValue "U_ColSum05", oRow, 0 '당월상환이자
            oDS_PH_PY315B.setValue "U_ColSum06", oRow, 0 '미상환잔액
            oDS_PH_PY315B.setValue "U_ColSum07", oRow, 0 '당월상환총액
            
            oMat1.LoadFromDataSource
        Else
            oDS_PH_PY315B.Offset = oRow - 1
            oDS_PH_PY315B.setValue "U_LineNum", oRow - 1, oRow
            oDS_PH_PY315B.setValue "U_ColReg01", oRow - 1, "" '부서
            oDS_PH_PY315B.setValue "U_ColReg02", oRow - 1, "" '직급
            oDS_PH_PY315B.setValue "U_ColReg03", oRow - 1, "" '사번
            oDS_PH_PY315B.setValue "U_ColReg04", oRow - 1, "" '성명
            oDS_PH_PY315B.setValue "U_ColDt01", oRow - 1, "" '대출일자
            oDS_PH_PY315B.setValue "U_ColSum01", oRow - 1, 0 '신규대출액
            oDS_PH_PY315B.setValue "U_ColQty01", oRow - 1, 0 '금리
            oDS_PH_PY315B.setValue "U_ColSum02", oRow - 1, 0 '대출총액
            oDS_PH_PY315B.setValue "U_ColSum03", oRow - 1, 0 '전월대출잔액
            oDS_PH_PY315B.setValue "U_ColSum04", oRow - 1, 0 '당월상환원금
            oDS_PH_PY315B.setValue "U_ColSum05", oRow - 1, 0 '당월상환이자
            oDS_PH_PY315B.setValue "U_ColSum06", oRow - 1, 0 '미상환잔액
            oDS_PH_PY315B.setValue "U_ColSum07", oRow - 1, 0 '당월상환총액
            
            oMat1.LoadFromDataSource
        End If
    ElseIf oMat1.VisualRowCount = 0 Then
        oDS_PH_PY315B.Offset = oRow
        oDS_PH_PY315B.setValue "U_LineNum", oRow, oRow + 1
        oDS_PH_PY315B.setValue "U_ColReg01", oRow, "" '부서
        oDS_PH_PY315B.setValue "U_ColReg02", oRow, "" '직급
        oDS_PH_PY315B.setValue "U_ColReg03", oRow, "" '사번
        oDS_PH_PY315B.setValue "U_ColReg04", oRow, "" '성명
        oDS_PH_PY315B.setValue "U_ColDt01", oRow, "" '대출일자
        oDS_PH_PY315B.setValue "U_ColSum01", oRow, 0 '신규대출액
        oDS_PH_PY315B.setValue "U_ColQty01", oRow, 0 '금리
        oDS_PH_PY315B.setValue "U_ColSum02", oRow, 0 '대출총액
        oDS_PH_PY315B.setValue "U_ColSum03", oRow, 0 '전월대출잔액
        oDS_PH_PY315B.setValue "U_ColSum04", oRow, 0 '당월상환원금
        oDS_PH_PY315B.setValue "U_ColSum05", oRow, 0 '당월상환이자
        oDS_PH_PY315B.setValue "U_ColSum06", oRow, 0 '미상환잔액
        oDS_PH_PY315B.setValue "U_ColSum07", oRow, 0 '당월상환총액
        
        oMat1.LoadFromDataSource
    End If
    
    Call oForm.Freeze(False)
    Exit Sub
PH_PY315_AddMatrixRow1_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY315_AddMatrixRow1_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY315_FormClear()
On Error GoTo PH_PY315_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY315'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PH_PY315_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY315_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY315_FormResize()
'******************************************************************************
'Function ID : PH_PY315_FormResize()
'해당모듈 : PH_PY315
'기능 : Form의 크기 변경 시 아이템들의 위치 및 크기 동적 변경
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PH_PY315_FormResize_Error

'    oForm.Items("Rect01").Height = oForm.Height - 180
'    oForm.Items("Rect01").Width = oForm.Width - 35

    Exit Sub

PH_PY315_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY315_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY315_DataValidCheck() As Boolean
On Error GoTo PH_PY315_DataValidCheck_Error
    PH_PY315_DataValidCheck = False
    Dim i       As Long
    Dim sQry    As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    '사업장
    If Trim(oForm.Items("CLTCOD").Specific.VALUE) = "" Then
        Sbo_Application.SetStatusBarMessage "사업장은 필수입니다.", bmt_Short, True
        oForm.Items("CLTCOD").CLICK ct_Regular
        PH_PY315_DataValidCheck = False
        Exit Function
    End If
    
    '상환일자
    If Trim(oForm.Items("RpmtDate").Specific.VALUE) = "" Then
        Sbo_Application.SetStatusBarMessage "상환일자는 필수입니다.", bmt_Short, True
        oForm.Items("RpmtDate").CLICK ct_Regular
        PH_PY315_DataValidCheck = False
        Exit Function
    End If
    
    '회차
'    If Trim(oForm.Items("Count").Specific.VALUE) = "" Then
'        Sbo_Application.SetStatusBarMessage "회차는 필수입니다.", bmt_Short, True
'        oForm.Items("Count").CLICK ct_Regular
'        PH_PY315_DataValidCheck = False
'        Exit Function
'    End If
    
    
'    '라인
'    If oMat1.VisualRowCount > 1 Then
'        For i = 1 To oMat1.VisualRowCount - 1
'
'            '학교
'            If oMat1.Columns("SchCls").Cells(i).Specific.Value = "" Then
'                Sbo_Application.SetStatusBarMessage "학교는 필수입니다.", bmt_Short, True
'                oMat1.Columns("SchCls").Cells(i).CLICK ct_Regular
'                PH_PY315_DataValidCheck = False
'                Exit Function
'            End If
'
'            '학교명
'            If oMat1.Columns("SchName").Cells(i).Specific.Value = "" Then
'                Sbo_Application.SetStatusBarMessage "학교명은 필수입니다.", bmt_Short, True
'                oMat1.Columns("SchName").Cells(i).CLICK ct_Regular
'                PH_PY315_DataValidCheck = False
'                Exit Function
'            End If
'
'            '학년
'            If oMat1.Columns("Grade").Cells(i).Specific.Value = "" Then
'                Sbo_Application.SetStatusBarMessage "학년은 필수입니다.", bmt_Short, True
'                oMat1.Columns("Grade").Cells(i).CLICK ct_Regular
'                PH_PY315_DataValidCheck = False
'                Exit Function
'            End If
'
'            '회차
'            If oMat1.Columns("Count").Cells(i).Specific.Value = "" Then
'                Sbo_Application.SetStatusBarMessage "회차는 필수입니다.", bmt_Short, True
'                oMat1.Columns("Count").Cells(i).CLICK ct_Regular
'                PH_PY315_DataValidCheck = False
'                Exit Function
'            End If
'
'        Next
'    Else
'        Sbo_Application.SetStatusBarMessage "라인 데이터가 없습니다.", bmt_Short, True
'        PH_PY315_DataValidCheck = False
'        Exit Function
'    End If
    
        oMat1.FlushToDataSource
    '// Matrix 마지막 행 삭제(DB 저장시)
    If oDS_PH_PY315B.Size > 1 Then oDS_PH_PY315B.RemoveRecord (oDS_PH_PY315B.Size - 1)
    
    oMat1.LoadFromDataSource
    
    PH_PY315_DataValidCheck = True
    
    Exit Function
    
    Set oRecordSet = Nothing
    
    
PH_PY315_DataValidCheck_Error:
    Set oRecordSet = Nothing
    PH_PY315_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY315_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY315_MTX01()

    '//메트릭스에 데이터 로드
    
    Dim i       As Long
    Dim sQry    As String
    
    Dim CLTCOD As String
    Dim RpmtDate As String
    Dim CntcCode As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY315_MTX01_Error

    Call oForm.Freeze(True)
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    CLTCOD = oForm.Items("CLTCOD").Specific.VALUE
    RpmtDate = oForm.Items("RpmtDate").Specific.VALUE
    CntcCode = oForm.Items("CntcCode").Specific.VALUE
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet.RecordCount, False)
    
    sQry = "EXEC PH_PY315_01 '" & CLTCOD & "','" & RpmtDate & "','" & CntcCode & "'"
    Call oRecordSet.DoQuery(sQry)
    
    oMat1.Clear
    oMat1.FlushToDataSource
    oMat1.LoadFromDataSource
    
    If (oRecordSet.RecordCount = 0) Then
        Call oMat1.Clear
        GoTo PH_PY315_MTX01_Exit
    End If
    
    For i = 0 To oRecordSet.RecordCount - 1
        If i <> 0 Then
            oDS_PH_PY315B.InsertRecord (i)
        End If
        oDS_PH_PY315B.Offset = i
        Call oDS_PH_PY315B.setValue("U_LineNum", i, i + 1)
        Call oDS_PH_PY315B.setValue("U_ColReg01", i, oRecordSet.Fields("TeamName").VALUE) '부서
        Call oDS_PH_PY315B.setValue("U_ColReg02", i, oRecordSet.Fields("JIGName").VALUE) '직급
        Call oDS_PH_PY315B.setValue("U_ColReg03", i, oRecordSet.Fields("CntcCode").VALUE) '사번
        Call oDS_PH_PY315B.setValue("U_ColReg04", i, oRecordSet.Fields("CntcName").VALUE) '성명
        Call oDS_PH_PY315B.setValue("U_ColDt01", i, Format(oRecordSet.Fields("LoanDate").VALUE, "YYYYMMDD")) '대출일자
        Call oDS_PH_PY315B.setValue("U_ColSum01", i, oRecordSet.Fields("NewLoanAmt").VALUE) '신규대출액
        Call oDS_PH_PY315B.setValue("U_ColQty01", i, oRecordSet.Fields("IntRate").VALUE) '금리
        Call oDS_PH_PY315B.setValue("U_ColSum02", i, oRecordSet.Fields("TotLoanAmt").VALUE) '대출총액
        Call oDS_PH_PY315B.setValue("U_ColSum03", i, oRecordSet.Fields("PrRmainAmt").VALUE) '전월대출잔액
        Call oDS_PH_PY315B.setValue("U_ColSum04", i, oRecordSet.Fields("RpmtAmt").VALUE) '당월상환원금
        Call oDS_PH_PY315B.setValue("U_ColSum05", i, oRecordSet.Fields("RpmtInt").VALUE) '당월상환이자
        Call oDS_PH_PY315B.setValue("U_ColSum06", i, oRecordSet.Fields("RmainAmt").VALUE) '미상환잔액
        Call oDS_PH_PY315B.setValue("U_ColSum07", i, oRecordSet.Fields("TotRpmtAmt").VALUE) '당월상환총액
        
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
PH_PY315_MTX01_Exit:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "W")
    Exit Sub
PH_PY315_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY315_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY315_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY315_Validate_Error
    PH_PY315_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY315A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", bmt_Short, True
        PH_PY315_Validate = False
        GoTo PH_PY315_Validate_Exit
    End If
'
    If ValidateType = "수정" Then

    ElseIf ValidateType = "행삭제" Then

    ElseIf ValidateType = "취소" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY315_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY315_Validate_Error:
    PH_PY315_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY315_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY315_Print_Report01()

    Dim DocNum          As String
    Dim ErrNum          As Integer
    Dim WinTitle        As String
    Dim ReportName      As String
    Dim sQry            As String
    Dim oRecordSet      As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY315_Print_Report01_Error
    
    Dim CLTCOD As String
    Dim RpmtDate As String
    Dim CntcCode As String
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    
     '/ ODBC 연결 체크
    If ConnectODBC = False Then
        GoTo PH_PY315_Print_Report01_Error
    End If
    
    
    '//인자 MOVE , Trim 시키기..
    CLTCOD = oForm.Items("CLTCOD").Specific.VALUE
    RpmtDate = oForm.Items("RpmtDate").Specific.VALUE
    CntcCode = oForm.Items("CntcCode").Specific.VALUE

    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/

    WinTitle = "[PH_PY315] 개인별 대부금 잔액현황"
    ReportName = "PH_PY315_01.rpt"
    
    '/ Formula 수식필드
    ReDim gRpt_Formula(0)
    ReDim gRpt_Formula_Value(0)
    
    
    '/ SubReport
    ReDim gRpt_SRptSqry(0)
    ReDim gRpt_SRptName(0)
    
'    ReDim gRpt_SFormula(0, 0)
'    ReDim gRpt_SFormula_Value(0, 0)
    
    
    '/ Procedure 실행"
    sQry = "EXEC PH_PY315_01 '" & CLTCOD & "','" & RpmtDate & "','" & CntcCode & "'"
    
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
        ErrNum = 1
        GoTo PH_PY315_Print_Report01_Error
    End If
    
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "1", "Y", "V", , 1) = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
    End If
    
    Set oRecordSet = Nothing
    Exit Sub
    
PH_PY315_Print_Report01_Error:
    If ErrNum = 1 Then
        Set oRecordSet = Nothing
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
    Else
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "PH_PY315_Print_Report01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    End If
    
End Sub



