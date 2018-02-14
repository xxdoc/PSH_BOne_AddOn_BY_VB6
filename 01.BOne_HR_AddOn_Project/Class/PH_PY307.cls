VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY307"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY307.cls
'//  Module         : 인사관리 > 기타
'//  Desc           : 학자금신청내역(분기별)
'//********************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form

Public oMat1 As SAPbouiCOM.Matrix
Public oMat2 As SAPbouiCOM.Matrix

Private oDS_PH_PY307A As SAPbouiCOM.DBDataSource
Private oDS_PH_PY307B As SAPbouiCOM.DBDataSource
Private oDS_PH_PY307C As SAPbouiCOM.DBDataSource

Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long

Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY307.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY307_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY307"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
'    oForm.DataBrowser.BrowseBy = "DocEntry"
        
    oForm.Freeze True
    Call PH_PY307_CreateItems
    Call PH_PY307_EnableMenus
    Call PH_PY307_SetDocument(oFromDocEntry01)
    Call PH_PY307_FormResize
    
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

Private Function PH_PY307_CreateItems() As Boolean

    Dim sQry        As String
    Dim i           As Long
    
    Dim oEdit       As SAPbouiCOM.EditText
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns    As SAPbouiCOM.Columns
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY307_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
'    Set oDS_PH_PY307A = oForm.DataSources.DBDataSources("@PH_PY307A")
    Set oDS_PH_PY307B = oForm.DataSources.DBDataSources("@PS_USERDS01")
'    Set oDS_PH_PY307C = oForm.DataSources.DBDataSources("@PS_USERDS02")
    
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
    
    '//부서
    Call oForm.DataSources.UserDataSources.Add("TeamCode", dt_SHORT_TEXT, 10)
    Set oCombo = oForm.Items("TeamCode").Specific
    oCombo.DataBind.SetBound True, "", "TeamCode"
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '1' AND U_UseYN= 'Y' AND U_Char2 = '" & oForm.Items.Item("CLTCOD").Specific.VALUE & "'"
    Call SetReDataCombo(oForm, sQry, oCombo, "Y")
    oForm.Items("TeamCode").DisplayDesc = True
    
    '분기_S
    Set oCombo = oForm.Items("Quarter").Specific
    oCombo.ValidValues.Add "", ""
    oCombo.ValidValues.Add "01", "1/4 혹은 1학기"
    oCombo.ValidValues.Add "02", "2/4"
    oCombo.ValidValues.Add "03", "3/4 혹은 2학기"
    oCombo.ValidValues.Add "04", "4/4"
    oCombo.Select 0, psk_Index
    oForm.Items("Quarter").DisplayDesc = True
    '분기_E
    
    '회차
    Set oCombo = oForm.Items("Count").Specific
    oCombo.ValidValues.Add "", ""
    oCombo.ValidValues.Add "01", "1차"
    oCombo.ValidValues.Add "02", "2차"
    oCombo.Select 0, psk_Index
    oForm.Items("Count").DisplayDesc = True
    
    '입학금포함여부
    Set oCombo = oForm.Items("EntFeeYN").Specific
'    oCombo.ValidValues.Add "", ""
    oCombo.ValidValues.Add "01", "포함"
    oCombo.ValidValues.Add "02", "제외"
    oCombo.Select 0, psk_Index
    oForm.Items("EntFeeYN").DisplayDesc = True
  
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
    
PH_PY307_CreateItems_Error:
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY307_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY307_EnableMenus()
    
    On Error GoTo PH_PY307_EnableMenus_Error
    
'    Call oForm.EnableMenu("1283", False)         '// 삭제
'    Call oForm.EnableMenu("1287", True)          '// 복제
''    Call oForm.EnableMenu("1286", True)         '// 닫기
'    Call oForm.EnableMenu("1284", True)         '// 취소
'    Call oForm.EnableMenu("1293", True)         '// 행삭제
    
    Exit Sub
    
PH_PY307_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY307_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY307_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY307_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY307_FormItemEnabled
'        Call PH_PY307_AddMatrixRow1
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY307_FormItemEnabled
        oForm.Items("DocEntry").Specific.VALUE = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY307_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY307_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY307_FormItemEnabled()
    On Error GoTo PH_PY307_FormItemEnabled_Error
    
    Dim oCombo As SAPbouiCOM.ComboBox
    Dim CLTCOD As String
        
    Call oForm.Freeze(True)
    If (oForm.Mode = fm_ADD_MODE) Then

        '폼 DocEntry 세팅
'        Call PH_PY307_FormClear
        
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        
        '년도 세팅
'        Call oDS_PH_PY307A.setValue("U_StdYear", 0, Format(Date, "YYYY"))
        oForm.Items("StdYear").Specific.VALUE = Format(Date, "YYYY")

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
    
PH_PY307_FormItemEnabled_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY307_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
'                        If PH_PY307_DataValidCheck = False Then
'                            BubbleEvent = False
'                        End If
'
'                        '//해야할일 작업
'                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
'                        If PH_PY307_DataValidCheck = False Then
'                            BubbleEvent = False
'                        End If
'                        '//해야할일 작업
'
'                    ElseIf oForm.Mode = fm_OK_MODE Then
'                    End If
'                End If
                If pval.ItemUID = "BtnSearch" Then
                
                    If oForm.Mode = fm_ADD_MODE Then
                        If PH_PY307_DataValidCheck = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
'
'                        '//해야할일 작업
                        Call PH_PY307_MTX01
'                        Call PH_PY307_MTX02
                        
                    End If
               
                End If
                
                If pval.ItemUID = "BtnPrint1" Then
                
                    If oForm.Mode = fm_ADD_MODE Then
                        If PH_PY307_DataValidCheck = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
'
'                        '//해야할일 작업
                        Call PH_PY307_Print_Report01
                        
                    End If
               
                End If
                
            ElseIf pval.BeforeAction = False Then
'                If pval.ItemUID = "1" Then
'                    If oForm.Mode = fm_ADD_MODE Then
'                        If pval.ActionSuccess = True Then
'                            Call PH_PY307_FormItemEnabled
'                            Call PH_PY307_AddMatrixRow1
'                        End If
'                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
'                        If pval.ActionSuccess = True Then
'                            Call PH_PY307_FormItemEnabled
'                            Call PH_PY307_AddMatrixRow1
'                        End If
'                    ElseIf oForm.Mode = fm_OK_MODE Then
'                        If pval.ActionSuccess = True Then
'                            Call PH_PY307_FormItemEnabled
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
                
                ElseIf pval.ItemUID = "SCode" And pval.CharPressed = "9" Then
                
                    If oForm.Items("SCode").Specific.VALUE = "" Then
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
                    Select Case pval.ItemUID
                           '//사업장이 바뀌면 부서와 담당 재설정
                           Case "CLTCOD"
                                '//부서
                                Set oCombo = oForm.Items("TeamCode").Specific
                                If oCombo.ValidValues.Count > 0 Then              '//삭제
                                   For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                                           oCombo.ValidValues.Remove i, psk_Index
                                   Next i
                                End If
                                '//현재 사업장으로 다시 Qry
                                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '1' AND U_UseYN= 'Y' AND U_Char2 = '" & oForm.Items.Item("CLTCOD").Specific.VALUE & "'"
                                Call SetReDataCombo(oForm, sQry, oCombo, "Y")
                                
'                                '//담당
'                                Set oCombo = oForm.Items("RspCode").Specific
'                                If oCombo.ValidValues.Count > 0 Then              '//삭제
'                                   For i = oCombo.ValidValues.Count - 1 To 0 Step -1
'                                           oCombo.ValidValues.Remove i, psk_Index
'                                   Next i
'                                End If
'                                '//현재 사업장으로 다시 Qry
'                                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '2' AND U_UseYN= 'Y' AND U_Char2 = '" & oForm.Items.Item("CLTCOD").Specific.VALUE & "'"
'                                Call SetReDataCombo(oForm, sQry, oCombo, "Y")
                                
'                           '//부서가 바뀌면 담당 재설정
'                           Case "TeamCode"
'                                '//담당은 그 부서의 담당만 표시
'                                Set oCombo = oForm.Items("RspCode").Specific
'                                If oCombo.ValidValues.Count > 0 Then              '//삭제
'                                   For i = oCombo.ValidValues.Count - 1 To 0 Step -1
'                                           oCombo.ValidValues.Remove i, psk_Index
'                                   Next i
'                                End If
'                                '//현재 사업장으로 다시 Qry
'                                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '2' AND U_UseYN= 'Y' AND U_Char1 = '" & oForm.Items.Item("TeamCode").Specific.VALUE & "' AND U_Char2 = '" & oForm.Items.Item("CLTCOD").Specific.VALUE & "'"
'                                Call SetReDataCombo(oForm, sQry, oCombo, "Y")
'
                    End Select
                    
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
                    
                        Case "SCode"
                            'Call oDS_PH_PY307A.setValue("U_CntcName", 0, MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & oForm.Items(pval.ItemUID).Specific.Value & "'"))
                        
                            oForm.Items("SName").Specific.VALUE = MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & oForm.Items(pval.ItemUID).Specific.VALUE & "'")
                        
                    End Select

                End If
                
            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_MATRIX_LOAD: '//11
            If pval.BeforeAction = True Then
            ElseIf pval.BeforeAction = False Then
                oMat1.LoadFromDataSource

                Call PH_PY307_FormItemEnabled
'                Call PH_PY307_AddMatrixRow1
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
                Set oDS_PH_PY307A = Nothing
                Set oDS_PH_PY307B = Nothing
                Set oDS_PH_PY307C = Nothing
                
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

                Call PH_PY307_FormResize
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
'                    Call MDC_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY307A", "Code")
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
                Call PH_PY307_AddMatrixRow1
                Call PH_PY307_AddMatrixRow2
            
        End Select
    ElseIf (pval.BeforeAction = False) Then
        Select Case pval.MenuUID
            Case "1283":
                oForm.Mode = fm_ADD_MODE
                Call PH_PY307_FormItemEnabled
                Call PH_PY307_AddMatrixRow1
            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//문서찾기
                Call PH_PY307_FormItemEnabled
                Call PH_PY307_AddMatrixRow1
                oForm.Items("DocEntry").CLICK ct_Regular
            Case "1282":    '//문서추가
                Call PH_PY307_FormItemEnabled
                Call PH_PY307_AddMatrixRow1
            Case "1288", "1289", "1290", "1291":
                Call PH_PY307_FormItemEnabled
            Case "1293" '// 행삭제

                 If oMat1.RowCount <> oMat1.VisualRowCount Then
                    oMat1.FlushToDataSource
    
                    While (i <= oDS_PH_PY307B.Size - 1)
                        If oDS_PH_PY307B.GetValue("U_LineNum", i) = "" Then
                            oDS_PH_PY307B.RemoveRecord (i)
                            i = 0
                        Else
                            i = i + 1
                        End If
                    Wend
                    
                    For i = 0 To oDS_PH_PY307B.Size
                        Call oDS_PH_PY307B.setValue("U_LineNum", i, i + 1)
                    Next i
                    
                    oMat1.LoadFromDataSource
                End If
                Call PH_PY307_AddMatrixRow1
                
            Case "7169": '엑셀 내보내기
            
                '엑셀 내보내기 이후 처리
                Call oForm.Freeze(True)
                Call oDS_PH_PY307B.RemoveRecord(oDS_PH_PY307B.Size - 1)
                Call oMat1.LoadFromDataSource
                
                Call oDS_PH_PY307C.RemoveRecord(oDS_PH_PY307C.Size - 1)
                Call oMat2.LoadFromDataSource
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

Sub PH_PY307_AddMatrixRow1()
    Dim oRow    As Long
    
    On Error GoTo PH_PY307_AddMatrixRow1_Error
    
    Call oForm.Freeze(True)
    
    '//[Mat1]
    oMat1.FlushToDataSource
    oRow = oMat1.VisualRowCount
    
    If oMat1.VisualRowCount > 0 Then
        If Trim(oDS_PH_PY307B.GetValue("U_LineNum", oRow - 1)) <> "" Then
            If oDS_PH_PY307B.Size <= oMat1.VisualRowCount Then
                oDS_PH_PY307B.InsertRecord (oRow)
            End If
            oDS_PH_PY307B.Offset = oRow
            oDS_PH_PY307B.setValue "U_LineNum", oRow, oRow + 1
            oDS_PH_PY307B.setValue "U_ColReg01", oRow, "" '신청인부서
            oDS_PH_PY307B.setValue "U_ColReg02", oRow, "" '신청인성명
            oDS_PH_PY307B.setValue "U_ColReg03", oRow, "" '자녀성명고등
            oDS_PH_PY307B.setValue "U_ColDt01", oRow, "" '자녀생년월일고등
            oDS_PH_PY307B.setValue "U_ColReg04", oRow, "" '자녀학년고등
            oDS_PH_PY307B.setValue "U_ColSum01", oRow, "" '자녀학자금고등
            oDS_PH_PY307B.setValue "U_ColReg05", oRow, "" '자녀성명전문대
            oDS_PH_PY307B.setValue "U_ColDt02", oRow, "" '자녀생년월일전문대
            oDS_PH_PY307B.setValue "U_ColReg06", oRow, "" '자녀학년전문대
            oDS_PH_PY307B.setValue "U_ColSum02", oRow, "" '자녀학자금전문대
            oDS_PH_PY307B.setValue "U_ColReg07", oRow, "" '회차
            oDS_PH_PY307B.setValue "U_ColReg08", oRow, "" '자녀성명대학
            oDS_PH_PY307B.setValue "U_ColDt03", oRow, "" '자녀생년월일대학
            oDS_PH_PY307B.setValue "U_ColReg09", oRow, "" '자녀학년대학
            oDS_PH_PY307B.setValue "U_ColSum03", oRow, "" '자녀학자금대학
            oDS_PH_PY307B.setValue "U_ColReg10", oRow, "" '회차
            oDS_PH_PY307B.setValue "U_ColReg04", oRow, "" '계
            
            oMat1.LoadFromDataSource
        Else
            oDS_PH_PY307B.Offset = oRow - 1
            oDS_PH_PY307B.setValue "U_LineNum", oRow - 1, oRow
            oDS_PH_PY307B.setValue "U_ColReg01", oRow - 1, "" '신청인부서
            oDS_PH_PY307B.setValue "U_ColReg02", oRow - 1, "" '신청인성명
            oDS_PH_PY307B.setValue "U_ColReg03", oRow - 1, "" '자녀성명고등
            oDS_PH_PY307B.setValue "U_ColDt01", oRow - 1, "" '자녀생년월일고등
            oDS_PH_PY307B.setValue "U_ColReg04", oRow - 1, "" '자녀학년고등
            oDS_PH_PY307B.setValue "U_ColSum01", oRow - 1, "" '자녀학자금고등
            oDS_PH_PY307B.setValue "U_ColReg05", oRow - 1, "" '자녀성명전문대
            oDS_PH_PY307B.setValue "U_ColDt02", oRow - 1, "" '자녀생년월일전문대
            oDS_PH_PY307B.setValue "U_ColReg06", oRow - 1, "" '자녀학년전문대
            oDS_PH_PY307B.setValue "U_ColSum02", oRow - 1, "" '자녀학자금전문대
            oDS_PH_PY307B.setValue "U_ColReg07", oRow - 1, "" '회차
            oDS_PH_PY307B.setValue "U_ColReg08", oRow - 1, "" '자녀성명대학
            oDS_PH_PY307B.setValue "U_ColDt03", oRow - 1, "" '자녀생년월일대학
            oDS_PH_PY307B.setValue "U_ColReg09", oRow - 1, "" '자녀학년대학
            oDS_PH_PY307B.setValue "U_ColSum03", oRow - 1, "" '자녀학자금대학
            oDS_PH_PY307B.setValue "U_ColReg10", oRow - 1, "" '회차
            oDS_PH_PY307B.setValue "U_ColReg04", oRow - 1, "" '계
            
            oMat1.LoadFromDataSource
        End If
    ElseIf oMat1.VisualRowCount = 0 Then
        oDS_PH_PY307B.Offset = oRow
        oDS_PH_PY307B.setValue "U_LineNum", oRow, oRow + 1
        oDS_PH_PY307B.setValue "U_ColReg01", oRow, "" '신청인부서
        oDS_PH_PY307B.setValue "U_ColReg02", oRow, "" '신청인성명
        oDS_PH_PY307B.setValue "U_ColReg03", oRow, "" '자녀성명고등
        oDS_PH_PY307B.setValue "U_ColDt01", oRow, "" '자녀생년월일고등
        oDS_PH_PY307B.setValue "U_ColReg04", oRow, "" '자녀학년고등
        oDS_PH_PY307B.setValue "U_ColSum01", oRow, "" '자녀학자금고등
        oDS_PH_PY307B.setValue "U_ColReg05", oRow, "" '자녀성명전문대
        oDS_PH_PY307B.setValue "U_ColDt02", oRow, "" '자녀생년월일전문대
        oDS_PH_PY307B.setValue "U_ColReg06", oRow, "" '자녀학년전문대
        oDS_PH_PY307B.setValue "U_ColSum02", oRow, "" '자녀학자금전문대
        oDS_PH_PY307B.setValue "U_ColReg07", oRow, "" '회차
        oDS_PH_PY307B.setValue "U_ColReg08", oRow, "" '자녀성명대학
        oDS_PH_PY307B.setValue "U_ColDt03", oRow, "" '자녀생년월일대학
        oDS_PH_PY307B.setValue "U_ColReg09", oRow, "" '자녀학년대학
        oDS_PH_PY307B.setValue "U_ColSum03", oRow, "" '자녀학자금대학
        oDS_PH_PY307B.setValue "U_ColReg10", oRow, "" '회차
        oDS_PH_PY307B.setValue "U_ColReg04", oRow, "" '계
        
        oMat1.LoadFromDataSource
    End If
    
    Call oForm.Freeze(False)
    Exit Sub
PH_PY307_AddMatrixRow1_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY307_AddMatrixRow1_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY307_AddMatrixRow2()
    Dim oRow    As Long
    
    On Error GoTo PH_PY307_AddMatrixRow2_Error
    
    Call oForm.Freeze(True)
    
    '//[Mat1]
    oMat2.FlushToDataSource
    oRow = oMat2.VisualRowCount
    
    If oMat2.VisualRowCount > 0 Then
        If Trim(oDS_PH_PY307C.GetValue("U_LineNum", oRow - 1)) <> "" Then
            If oDS_PH_PY307C.Size <= oMat2.VisualRowCount Then
                oDS_PH_PY307C.InsertRecord (oRow)
            End If
            oDS_PH_PY307C.Offset = oRow
            oDS_PH_PY307C.setValue "U_LineNum", oRow, oRow + 1
            oDS_PH_PY307C.setValue "U_ColReg01", oRow, "" '부서/담당명
            oDS_PH_PY307C.setValue "U_ColReg02", oRow, "" '고등인원
            oDS_PH_PY307C.setValue "U_ColSum01", oRow, "" '고등학자금
            oDS_PH_PY307C.setValue "U_ColReg03", oRow, "" '전문대인원
            oDS_PH_PY307C.setValue "U_ColSum02", oRow, "" '전문대학자금
            oDS_PH_PY307C.setValue "U_ColReg04", oRow, "" '대학인원
            oDS_PH_PY307C.setValue "U_ColSum03", oRow, "" '대학학자금
            oDS_PH_PY307C.setValue "U_ColSum04", oRow, "" '계
            
            oMat2.LoadFromDataSource
        Else
            oDS_PH_PY307C.Offset = oRow - 1
            oDS_PH_PY307C.setValue "U_LineNum", oRow - 1, oRow
            oDS_PH_PY307C.setValue "U_ColReg01", oRow - 1, "" '부서/담당명
            oDS_PH_PY307C.setValue "U_ColReg02", oRow - 1, "" '고등인원
            oDS_PH_PY307C.setValue "U_ColSum01", oRow - 1, "" '고등학자금
            oDS_PH_PY307C.setValue "U_ColReg03", oRow - 1, "" '전문대인원
            oDS_PH_PY307C.setValue "U_ColSum02", oRow - 1, "" '전문대학자금
            oDS_PH_PY307C.setValue "U_ColReg04", oRow - 1, "" '대학인원
            oDS_PH_PY307C.setValue "U_ColSum03", oRow - 1, "" '대학학자금
            oDS_PH_PY307C.setValue "U_ColSum04", oRow - 1, "" '계
            
            oMat2.LoadFromDataSource
        End If
    ElseIf oMat2.VisualRowCount = 0 Then
        oDS_PH_PY307C.Offset = oRow
        oDS_PH_PY307C.setValue "U_LineNum", oRow, oRow + 1
        oDS_PH_PY307C.setValue "U_ColReg01", oRow, "" '부서/담당명
        oDS_PH_PY307C.setValue "U_ColReg02", oRow, "" '고등인원
        oDS_PH_PY307C.setValue "U_ColSum01", oRow, "" '고등학자금
        oDS_PH_PY307C.setValue "U_ColReg03", oRow, "" '전문대인원
        oDS_PH_PY307C.setValue "U_ColSum02", oRow, "" '전문대학자금
        oDS_PH_PY307C.setValue "U_ColReg04", oRow, "" '대학인원
        oDS_PH_PY307C.setValue "U_ColSum03", oRow, "" '대학학자금
        oDS_PH_PY307C.setValue "U_ColSum04", oRow, "" '계
        
        oMat2.LoadFromDataSource
    End If
    
    Call oForm.Freeze(False)
    Exit Sub
PH_PY307_AddMatrixRow2_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY307_AddMatrixRow2_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY307_FormClear()
On Error GoTo PH_PY307_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY307'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PH_PY307_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY307_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY307_FormResize()
'******************************************************************************
'Function ID : PH_PY307_FormResize()
'해당모듈 : PH_PY307
'기능 : Form의 크기 변경 시 아이템들의 위치 및 크기 동적 변경
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PH_PY307_FormResize_Error

'    oForm.Items("Rect01").Height = oForm.Height - 180
'    oForm.Items("Rect01").Width = oForm.Width - 35

    Exit Sub

PH_PY307_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY307_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY307_DataValidCheck() As Boolean
On Error GoTo PH_PY307_DataValidCheck_Error
    PH_PY307_DataValidCheck = False
    Dim i       As Long
    Dim sQry    As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    '사업장
    If Trim(oForm.Items("CLTCOD").Specific.VALUE) = "" Then
        Sbo_Application.SetStatusBarMessage "사업장은 필수입니다.", bmt_Short, True
        oForm.Items("CLTCOD").CLICK ct_Regular
        PH_PY307_DataValidCheck = False
        Exit Function
    End If
    
    '년도
    If Trim(oForm.Items("StdYear").Specific.VALUE) = "" Then
        Sbo_Application.SetStatusBarMessage "년도는 필수입니다.", bmt_Short, True
        oForm.Items("StdYear").CLICK ct_Regular
        PH_PY307_DataValidCheck = False
        Exit Function
    End If

    '분기
    If Trim(oForm.Items("Quarter").Specific.VALUE) = "" Then
        Sbo_Application.SetStatusBarMessage "분기는 필수입니다.", bmt_Short, True
        oForm.Items("Quarter").CLICK ct_Regular
        PH_PY307_DataValidCheck = False
        Exit Function
    End If
    
    '회차
'    If Trim(oForm.Items("Count").Specific.VALUE) = "" Then
'        Sbo_Application.SetStatusBarMessage "회차는 필수입니다.", bmt_Short, True
'        oForm.Items("Count").CLICK ct_Regular
'        PH_PY307_DataValidCheck = False
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
'                PH_PY307_DataValidCheck = False
'                Exit Function
'            End If
'
'            '학교명
'            If oMat1.Columns("SchName").Cells(i).Specific.Value = "" Then
'                Sbo_Application.SetStatusBarMessage "학교명은 필수입니다.", bmt_Short, True
'                oMat1.Columns("SchName").Cells(i).CLICK ct_Regular
'                PH_PY307_DataValidCheck = False
'                Exit Function
'            End If
'
'            '학년
'            If oMat1.Columns("Grade").Cells(i).Specific.Value = "" Then
'                Sbo_Application.SetStatusBarMessage "학년은 필수입니다.", bmt_Short, True
'                oMat1.Columns("Grade").Cells(i).CLICK ct_Regular
'                PH_PY307_DataValidCheck = False
'                Exit Function
'            End If
'
'            '회차
'            If oMat1.Columns("Count").Cells(i).Specific.Value = "" Then
'                Sbo_Application.SetStatusBarMessage "회차는 필수입니다.", bmt_Short, True
'                oMat1.Columns("Count").Cells(i).CLICK ct_Regular
'                PH_PY307_DataValidCheck = False
'                Exit Function
'            End If
'
'        Next
'    Else
'        Sbo_Application.SetStatusBarMessage "라인 데이터가 없습니다.", bmt_Short, True
'        PH_PY307_DataValidCheck = False
'        Exit Function
'    End If
    
        oMat1.FlushToDataSource
    '// Matrix 마지막 행 삭제(DB 저장시)
    If oDS_PH_PY307B.Size > 1 Then oDS_PH_PY307B.RemoveRecord (oDS_PH_PY307B.Size - 1)
    
    oMat1.LoadFromDataSource
    
    PH_PY307_DataValidCheck = True
    
    Exit Function
    
    Set oRecordSet = Nothing
    
    
PH_PY307_DataValidCheck_Error:
    Set oRecordSet = Nothing
    PH_PY307_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY307_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY307_MTX01()

    '//메트릭스에 데이터 로드
    
    Dim i       As Long
    Dim sQry    As String
    
    Dim Param01 As String
    Dim Param02 As String
    Dim Param03 As String
    Dim Param04 As String
    Dim Param05 As String
    Dim Param06 As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY307_MTX01_Error

    Call oForm.Freeze(True)
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    Param01 = Trim(oForm.Items("CLTCOD").Specific.VALUE)
    Param02 = Trim(oForm.Items("StdYear").Specific.VALUE)
    Param03 = Trim(oForm.Items("Quarter").Specific.VALUE)
    Param04 = Trim(oForm.Items("Count").Specific.VALUE)
    Param05 = Trim(oForm.Items("TeamCode").Specific.VALUE)
    Param06 = Trim(oForm.Items("EntFeeYN").Specific.VALUE)
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet.RecordCount, False)
    
    sQry = "EXEC PH_PY307_01 '" & Param01 & "','" & Param02 & "','" & Param03 & "','" & Param04 & "','" & Param05 & "','" & Param06 & "'"
    Call oRecordSet.DoQuery(sQry)
    
    oMat1.Clear
    oMat1.FlushToDataSource
    oMat1.LoadFromDataSource
    
    If (oRecordSet.RecordCount = 0) Then
        Call oMat1.Clear
        GoTo PH_PY307_MTX01_Exit
    End If
    
    For i = 0 To oRecordSet.RecordCount - 1
        If i <> 0 Then
            oDS_PH_PY307B.InsertRecord (i)
        End If
        oDS_PH_PY307B.Offset = i
        Call oDS_PH_PY307B.setValue("U_LineNum", i, i + 1)
        Call oDS_PH_PY307B.setValue("U_ColReg01", i, oRecordSet.Fields("Team").VALUE) '신청인부서
        Call oDS_PH_PY307B.setValue("U_ColReg02", i, oRecordSet.Fields("Rsp").VALUE) '담당
        Call oDS_PH_PY307B.setValue("U_ColReg16", i, oRecordSet.Fields("Cls").VALUE) '반추가(2017.02.06 송명규, 최용석 요청)
        Call oDS_PH_PY307B.setValue("U_ColReg03", i, oRecordSet.Fields("CntcCode").VALUE) '사번
        Call oDS_PH_PY307B.setValue("U_ColReg04", i, oRecordSet.Fields("CntcName").VALUE) '신청인성명
        Call oDS_PH_PY307B.setValue("U_ColReg05", i, oRecordSet.Fields("Name1").VALUE) '자녀성명고등
        Call oDS_PH_PY307B.setValue("U_ColReg06", i, oRecordSet.Fields("BirDat1").VALUE) '자녀생년월일고등
        Call oDS_PH_PY307B.setValue("U_ColReg07", i, oRecordSet.Fields("Grade1").VALUE) '자녀학년고등
        Call oDS_PH_PY307B.setValue("U_ColSum01", i, oRecordSet.Fields("Amt1").VALUE) '자녀학자금고등
        Call oDS_PH_PY307B.setValue("U_ColReg08", i, oRecordSet.Fields("Name2").VALUE) '자녀성명전문대
        Call oDS_PH_PY307B.setValue("U_ColReg09", i, oRecordSet.Fields("BirDat2").VALUE) '자녀생년월일전문대
        Call oDS_PH_PY307B.setValue("U_ColReg10", i, oRecordSet.Fields("Grade2").VALUE) '자녀학년전문대
        Call oDS_PH_PY307B.setValue("U_ColSum02", i, oRecordSet.Fields("Amt2").VALUE) '자녀학자금전문대
        Call oDS_PH_PY307B.setValue("U_ColReg11", i, oRecordSet.Fields("Count2").VALUE) '회차
        Call oDS_PH_PY307B.setValue("U_ColReg12", i, oRecordSet.Fields("Name3").VALUE) '자녀성명대학
        Call oDS_PH_PY307B.setValue("U_ColReg13", i, oRecordSet.Fields("BirDat3").VALUE) '자녀생년월일대학
        Call oDS_PH_PY307B.setValue("U_ColReg14", i, oRecordSet.Fields("Grade3").VALUE) '자녀학년대학
        Call oDS_PH_PY307B.setValue("U_ColSum03", i, oRecordSet.Fields("Amt3").VALUE) '자녀학자금대학
        Call oDS_PH_PY307B.setValue("U_ColReg15", i, oRecordSet.Fields("Count3").VALUE) '회차
'        Call oDS_PH_PY307B.setValue("U_ColSum04", i, oRecordSet.Fields("Total").VALUE) '계
        
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
PH_PY307_MTX01_Exit:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "W")
    Exit Sub
PH_PY307_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY307_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY307_MTX02()

    '//메트릭스에 데이터 로드
    
    Dim i       As Long
    Dim sQry    As String
    
    Dim Param01 As String
    Dim Param02 As String
    Dim Param03 As String
    Dim Param04 As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY307_MTX02_Error

    Call oForm.Freeze(True)
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    Param01 = oForm.Items("CLTCOD").Specific.VALUE
    Param02 = oForm.Items("StdYear").Specific.VALUE
    Param03 = oForm.Items("Quarter").Specific.VALUE
    Param04 = oForm.Items("Count").Specific.VALUE
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet.RecordCount, False)
    
    sQry = "EXEC PH_PY307_02 '" & Param01 & "','" & Param02 & "','" & Param03 & "','" & Param04 & "'"
    Call oRecordSet.DoQuery(sQry)
    
    oMat2.Clear
    oMat2.FlushToDataSource
    oMat2.LoadFromDataSource
    
    If (oRecordSet.RecordCount = 0) Then
        Call oMat2.Clear
        GoTo PH_PY307_MTX02_Exit
    End If
    
    For i = 0 To oRecordSet.RecordCount - 1
        If i <> 0 Then
            oDS_PH_PY307C.InsertRecord (i)
        End If
        oDS_PH_PY307C.Offset = i
        Call oDS_PH_PY307C.setValue("U_LineNum", i, i + 1)
        Call oDS_PH_PY307C.setValue("U_ColReg01", i, oRecordSet.Fields("Team").VALUE) '부서/담당명
        Call oDS_PH_PY307C.setValue("U_ColReg02", i, oRecordSet.Fields("ManCnt1").VALUE) '고등인원
        Call oDS_PH_PY307C.setValue("U_ColSum01", i, oRecordSet.Fields("Amt1").VALUE) '고등학자금
        Call oDS_PH_PY307C.setValue("U_ColReg03", i, oRecordSet.Fields("ManCnt2").VALUE) '전문대인원
        Call oDS_PH_PY307C.setValue("U_ColSum02", i, oRecordSet.Fields("Amt2").VALUE) '전문대학자금
        Call oDS_PH_PY307C.setValue("U_ColReg04", i, oRecordSet.Fields("ManCnt3").VALUE) '대학인원
        Call oDS_PH_PY307C.setValue("U_ColSum03", i, oRecordSet.Fields("Amt3").VALUE) '대학학자금
        Call oDS_PH_PY307C.setValue("U_ColSum04", i, oRecordSet.Fields("Total").VALUE) '계
        
        oRecordSet.MoveNext
        ProgressBar01.VALUE = ProgressBar01.VALUE + 1
        ProgressBar01.Text = ProgressBar01.VALUE & "/" & oRecordSet.RecordCount & "건 조회중...!"
        
    Next
    
    oMat2.LoadFromDataSource
    oMat2.AutoResizeColumns
    oForm.Update
    
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Sub
PH_PY307_MTX02_Exit:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "W")
    Exit Sub
PH_PY307_MTX02_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY307_MTX02_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY307_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY307_Validate_Error
    PH_PY307_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY307A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", bmt_Short, True
        PH_PY307_Validate = False
        GoTo PH_PY307_Validate_Exit
    End If
'
    If ValidateType = "수정" Then

    ElseIf ValidateType = "행삭제" Then

    ElseIf ValidateType = "취소" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY307_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY307_Validate_Error:
    PH_PY307_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY307_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY307_Print_Report01()

    Dim DocNum          As String
    Dim ErrNum          As Integer
    Dim WinTitle        As String
    Dim ReportName      As String
    Dim sQry            As String
    Dim oRecordSet      As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY307_Print_Report01_Error
    
    Dim Param01 As String
    Dim Param02 As String
    Dim Param03 As String
    Dim Param04 As String
    Dim Param05 As String
    Dim Param06 As String
    
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    
     '/ ODBC 연결 체크
    If ConnectODBC = False Then
        GoTo PH_PY307_Print_Report01_Error
    End If
    
    
    '//인자 MOVE , Trim 시키기..
    Param01 = oForm.Items("CLTCOD").Specific.VALUE
    Param02 = oForm.Items("StdYear").Specific.VALUE
    Param03 = oForm.Items("Quarter").Specific.VALUE
    Param04 = oForm.Items("Count").Specific.VALUE
    Param05 = oForm.Items("TeamCode").Specific.VALUE
    Param06 = oForm.Items("EntFeeYN").Specific.VALUE
    
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/

    WinTitle = "[PH_PY307] 학자금신청내역(분기별)"
    ReportName = "PH_PY307_01.rpt"
    
    '/ Formula 수식필드
    ReDim gRpt_Formula(4)
    ReDim gRpt_Formula_Value(4)
    
    gRpt_Formula(1) = "CLTCOD" '사업장
    sQry = "SELECT U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y' AND U_Code = '" & Param01 & "'"
    Call oRecordSet.DoQuery(sQry)
    gRpt_Formula_Value(1) = oRecordSet.Fields(0).VALUE
    
    gRpt_Formula(2) = "YYYY" '년도
    gRpt_Formula_Value(2) = Param02 'oRecordSet.Fields(0).VALUE
    
    gRpt_Formula(3) = "BUNGI" '분기
    If Param03 = "01" Then
        gRpt_Formula_Value(3) = "1/4 혹은 1학기"
    ElseIf Param03 = "02" Then
        gRpt_Formula_Value(3) = "2/4"
    ElseIf Param03 = "03" Then
        gRpt_Formula_Value(3) = "3/4 혹은 2학기"
    ElseIf Param03 = "04" Then
        gRpt_Formula_Value(3) = "4/4"
    End If
    
    gRpt_Formula(4) = "CHASU" '회차
    If Param04 = "01" Then
        gRpt_Formula_Value(4) = "1차"
    ElseIf Param04 = "02" Then
        gRpt_Formula_Value(4) = "2차"
    Else
        gRpt_Formula_Value(4) = "전체"
    End If
    
    '/ SubReport
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)
    
    
    '/ Procedure 실행"
    sQry = "EXEC PH_PY307_01 '" & Param01 & "','" & Param02 & "','" & Param03 & "','" & Param04 & "','" & Param05 & "','" & Param06 & "'"
    
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
        ErrNum = 1
        GoTo PH_PY307_Print_Report01_Error
    End If
    
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "1", "Y", "V", , 2) = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
    End If
    
    Set oRecordSet = Nothing
    Exit Sub
    
PH_PY307_Print_Report01_Error:
    If ErrNum = 1 Then
        Set oRecordSet = Nothing
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
    Else
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "PH_PY307_Print_Report01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    End If
    
End Sub

