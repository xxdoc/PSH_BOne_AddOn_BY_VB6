VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PYA60"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PYA60.cls
'//  Module         : PH
'//  Desc           : 학자금신청내역(집계)
'//  작성자         : NGY
'//  DATE           : 2016.01.20
'//********************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form

''// 그리드 사용시
'Public oGrid1           As SAPbouiCOM.Grid
'Public oDS_PH_PYA60     As SAPbouiCOM.DataTable
'
''// 매트릭스 사용시
'Public oMat1 As SAPbouiCOM.Matrix
'Private oDS_PH_PYA60A As SAPbouiCOM.DBDataSource
'Private oDS_PH_PYA60B As SAPbouiCOM.DBDataSource

Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long

Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PYA60.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PYA60_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PYA60"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
'    oForm.DataBrowser.BrowseBy = "Code"
        
    oForm.Freeze True
    Call PH_PYA60_CreateItems
    Call PH_PYA60_EnableMenus
    Call PH_PYA60_SetDocument(oFromDocEntry01)
'    Call PH_PYA60_FormResize
    
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

Private Function PH_PYA60_CreateItems() As Boolean

    Dim sQry        As String
    Dim i           As Long
    
    Dim oCheck      As SAPbouiCOM.CheckBox
    Dim oEdit       As SAPbouiCOM.EditText
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns    As SAPbouiCOM.Columns
    Dim optBtn      As SAPbouiCOM.OptionBtn
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PYA60_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    
    '//----------------------------------------------------------------------------------------------
    '// 아이템 설정
    '//----------------------------------------------------------------------------------------------
    
    '//사업장
    Call oForm.DataSources.UserDataSources.Add("CLTCOD", dt_SHORT_TEXT, 10)
    Set oCombo = oForm.Items("CLTCOD").Specific

    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
    oRecordSet.DoQuery sQry
    Call oCombo.ValidValues.Add("%", "전 사업장")
    Do Until oRecordSet.EOF
        oCombo.ValidValues.Add Trim(oRecordSet.Fields(0).VALUE), Trim(oRecordSet.Fields(1).VALUE)
        oRecordSet.MoveNext
    Loop
    Call oCombo.Select(0, psk_Index)
    
    
    '// 접속자에 따른 권한별 사업장 콤보박스세팅
    'Call CLTCOD_Select(oForm, "CLTCOD")
    
    'oForm.Items("StdYear").Specific.String = Format(Now, "YYYY") - 1
    oForm.Items("StdYear").Specific.String = Format(Now, "YYYY")
    
    '분기_S
    Set oCombo = oForm.Items("Quarter").Specific
    oCombo.ValidValues.Add "%", "전체"
    oCombo.ValidValues.Add "01", "1/4 혹은 1학기"
    oCombo.ValidValues.Add "02", "2/4"
    oCombo.ValidValues.Add "03", "3/4 혹은 2학기"
    oCombo.ValidValues.Add "04", "4/4"
    oCombo.Select 0, psk_Index
    oForm.Items("Quarter").DisplayDesc = True
    '분기_E
    
    '회차
    Set oCombo = oForm.Items("Count").Specific
    oCombo.ValidValues.Add "%", "전체"
    oCombo.ValidValues.Add "01", "1차"
    oCombo.ValidValues.Add "02", "2차"
    oCombo.Select 0, psk_Index
    oForm.Items("Count").DisplayDesc = True
    
    '입학금포함여부
    Set oCombo = oForm.Items("EntFeeYN").Specific
    oCombo.ValidValues.Add "01", "포함"
    oCombo.ValidValues.Add "02", "제외"
    oCombo.Select 0, psk_Index
    oForm.Items("EntFeeYN").DisplayDesc = True
    
    '출력선택 RAD
    Call oForm.DataSources.UserDataSources.Add("OptionDS", dt_SHORT_TEXT, 1)
    Set optBtn = oForm.Items("Rad01").Specific
    optBtn.ValOn = "1"
    optBtn.ValOff = "0"
    optBtn.DataBind.SetBound True, , "OptionDS"
    
    optBtn.Selected = True
    
    Set optBtn = oForm.Items("Rad02").Specific
    optBtn.ValOn = "2"
    optBtn.ValOff = "0"
    optBtn.DataBind.SetBound True, , "OptionDS"
    optBtn.GroupWith ("Rad01")
    
    '//커서를 첫번째 ITEM으로 지정
    oForm.ActiveItem = "CLTCOD"
    
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
    
PH_PYA60_CreateItems_Error:
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PYA60_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PYA60_EnableMenus()
    
    On Error GoTo PH_PYA60_EnableMenus_Error

    Call oForm.EnableMenu("1283", True)     '//제거
    Call oForm.EnableMenu("1284", False)    '//취소
    Call oForm.EnableMenu("1293", True)     '//행삭제
    
    Exit Sub
    
PH_PYA60_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PYA60_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PYA60_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PYA60_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PYA60_FormItemEnabled
        Call PH_PYA60_AddMatrixRow
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PYA60_FormItemEnabled
        oForm.Items("Code").Specific.VALUE = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PYA60_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PYA60_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PYA60_FormItemEnabled()
    Dim oCombo      As SAPbouiCOM.ComboBox
    
    On Error GoTo PH_PYA60_FormItemEnabled_Error
    
    
    Call oForm.Freeze(True)
    If (oForm.Mode = fm_ADD_MODE) Then

        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", False)    '//문서추가
        
    ElseIf (oForm.Mode = fm_FIND_MODE) Then

        Call oForm.EnableMenu("1281", False)    '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가
    ElseIf (oForm.Mode = fm_OK_MODE) Then

        
        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가

    End If
    Call oForm.Freeze(False)
    Exit Sub
    
PH_PYA60_FormItemEnabled_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PYA60_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
                If pval.ItemUID = "Btn01" Then
                    Call PH_PYA60_DataValidCheck
                    Call PH_PYA60_Print_Report01
                End If
            ElseIf pval.BeforeAction = False Then


            End If
     
'----------------------------------------------------------
        Case et_KEY_DOWN: '//2

'----------------------------------------------------------
        Case et_GOT_FOCUS: '//3
            Select Case pval.ItemUID
            Case "Mat01", "Grid01"
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
 
                           
                    End Select
                End If
            End If
            
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_CLICK: '//6
            If pval.BeforeAction = True Then
                Select Case pval.ItemUID
                Case "Mat01"
'                    If pval.Row > 0 Then
'                        Call oMat1.SelectRow(pval.Row, True, False)
'                    End If
                End Select
                
                Select Case pval.ItemUID
                Case "Mat01", "Grid01"
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
                    Select Case pval.ItemUID
                      
                        
                    End Select
                End If
            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_MATRIX_LOAD: '//11
            If pval.BeforeAction = True Then
            ElseIf pval.BeforeAction = False Then
               'oMat1.LoadFromDataSource

                Call PH_PYA60_FormItemEnabled
                Call PH_PYA60_AddMatrixRow
                
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
'                Set oDS_PH_PYA60A = Nothing
'                Set oDS_PH_PYA60B = Nothing
                
                'Set oMat1 = Nothing
                'Set oGrid1 = Nothing
                
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
        End Select
    ElseIf (pval.BeforeAction = False) Then
        Select Case pval.MenuUID
            Case "1283":
                oForm.Mode = fm_ADD_MODE
                Call PH_PYA60_FormItemEnabled
                Call PH_PYA60_AddMatrixRow
            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//문서찾기
                Call PH_PYA60_FormItemEnabled
                Call PH_PYA60_AddMatrixRow
                oForm.Items("Code").CLICK ct_Regular
            Case "1282":    '//문서추가
                Call PH_PYA60_FormItemEnabled
                Call PH_PYA60_AddMatrixRow
            Case "1288", "1289", "1290", "1291":
                Call PH_PYA60_FormItemEnabled
            Case "1293" '// 행삭제
'                '// [MAT1 용]
'                 If oMat1.RowCount <> oMat1.VisualRowCount Then
'                    oMat1.FlushToDataSource
'
'                    While (i <= oDS_PH_PYA60B.Size - 1)
'                        If oDS_PH_PYA60B.GetValue("U_FILD01", i) = "" Then
'                            oDS_PH_PYA60B.RemoveRecord (i)
'                            i = 0
'                        Else
'                            i = i + 1
'                        End If
'                    Wend
'
'                    For i = 0 To oDS_PH_PYA60B.Size
'                        Call oDS_PH_PYA60B.setValue("U_LineNum", i, i + 1)
'                    Next i
'
'                    oMat1.LoadFromDataSource
                'End If
                Call PH_PYA60_AddMatrixRow
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

Sub PH_PYA60_AddMatrixRow()
    Dim oRow    As Long
    
    On Error GoTo PH_PYA60_AddMatrixRow_Error
    
    Call oForm.Freeze(True)
    
  
    
    Call oForm.Freeze(False)
    Exit Sub
PH_PYA60_AddMatrixRow_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PYA60_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PYA60_FormClear()
On Error GoTo PH_PYA60_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PYA60'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PH_PYA60_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PYA60_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PYA60_DataValidCheck() As Boolean
On Error GoTo PH_PYA60_DataValidCheck_Error
    PH_PYA60_DataValidCheck = False
    Dim i       As Long
    Dim sQry    As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '//----------------------------------------------------------------------------------
    '//필수 체크
    '//----------------------------------------------------------------------------------

    
    PH_PYA60_DataValidCheck = True
    
    Exit Function
    
    Set oRecordSet = Nothing
    
    
PH_PYA60_DataValidCheck_Error:
    Set oRecordSet = Nothing
    PH_PYA60_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PYA60_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Function PH_PYA60_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PYA60_Validate_Error
    PH_PYA60_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PYA60A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", bmt_Short, True
        PH_PYA60_Validate = False
        GoTo PH_PYA60_Validate_Exit
    End If
'
    If ValidateType = "수정" Then

    ElseIf ValidateType = "행삭제" Then

    ElseIf ValidateType = "취소" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PYA60_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PYA60_Validate_Error:
    PH_PYA60_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PYA60_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Private Sub PH_PYA60_Print_Report01()

    Dim DocNum          As String
    Dim ErrNum          As Integer
    Dim WinTitle        As String
    Dim ReportName      As String
    Dim sQry            As String
    Dim oRecordSet      As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PYA60_Print_Report01_Error
    
    Dim CLTCOD          As String
    Dim yyyy            As String
    Dim StdYear         As String
    Dim Quarter         As String
    Dim Count           As String
    Dim EntFeeYN        As String
    Dim OptBtnValue     As String
   
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
   
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

     '/ ODBC 연결 체크
    If ConnectODBC = False Then
        GoTo PH_PYA60_Print_Report01_Error
    End If
    
    '//인자 MOVE , Trim 시키기..
    CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE)
    yyyy = Trim(oForm.Items("StdYear").Specific.VALUE)
    Quarter = Trim(oForm.Items("Quarter").Specific.VALUE)
    Count = Trim(oForm.Items("Count").Specific.VALUE)
    EntFeeYN = Trim(oForm.Items("EntFeeYN").Specific.VALUE)
    OptBtnValue = oForm.DataSources.UserDataSources("OptionDS").VALUE
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/

    WinTitle = "[PH_PYA60] 학자금신청내역(집계)"
    
    If OptBtnValue = "1" Then
       ReportName = "PH_PYA60_01.rpt"
    Else
       ReportName = "PH_PYA60_02.rpt"
    End If
    
    '/ Formula 수식필드
    ReDim gRpt_Formula(4)
    ReDim gRpt_Formula_Value(4)
    
    gRpt_Formula(1) = "CLTCOD"
    sQry = "SELECT U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y' AND U_Code = '" & CLTCOD & "'"
    Call oRecordSet.DoQuery(sQry)
    If CLTCOD = "%" Then
       gRpt_Formula_Value(1) = "전 사업장"
    Else
       gRpt_Formula_Value(1) = oRecordSet.Fields(0).VALUE
    End If
    
    gRpt_Formula(2) = "YYYY"
    gRpt_Formula_Value(2) = yyyy

    gRpt_Formula(3) = "BUNGI" '분기
    If Quarter = "01" Then
        gRpt_Formula_Value(3) = "1/4 혹은 1학기"
    ElseIf Quarter = "02" Then
        gRpt_Formula_Value(3) = "2/4"
    ElseIf Quarter = "03" Then
        gRpt_Formula_Value(3) = "3/4 혹은 2학기"
    ElseIf Quarter = "04" Then
        gRpt_Formula_Value(3) = "4/4"
    Else
        gRpt_Formula_Value(3) = "전체"
    End If
    
    gRpt_Formula(4) = "CHASU" '회차
    If Count = "01" Then
        gRpt_Formula_Value(4) = "1차"
    ElseIf Count = "02" Then
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
    
    If OptBtnValue = "1" Then
        sQry = "EXEC [PH_PYA60_01] '" & CLTCOD & "', '" & yyyy & "', '" & Quarter & "', '" & Count & "', '" & EntFeeYN & "'"
    Else
        sQry = "EXEC [PH_PYA60_02] '" & CLTCOD & "', '" & yyyy & "', '" & Quarter & "', '" & Count & "', '" & EntFeeYN & "'"
    End If
    
    
    
'    oRecordSet.DoQuery sQry
'    If oRecordSet.RecordCount = 0 Then
'        ErrNum = 1
'        GoTo PH_PYA60_Print_Report01_Error
'    End If
    
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "1", "Y", "V", , 1) = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
    End If
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set oRecordSet = Nothing
    Exit Sub
    
PH_PYA60_Print_Report01_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    Set oRecordSet = Nothing

    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
    Else
        Sbo_Application.SetStatusBarMessage "PH_PYA60_Print_Report01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    End If
    
End Sub





























