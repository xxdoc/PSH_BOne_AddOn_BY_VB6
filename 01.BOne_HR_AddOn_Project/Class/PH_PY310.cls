VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY310"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY310.cls
'//  Module         : 인사관리 > 기타 > 대부금관리
'//  Desc           : 대부금개별상환
'//********************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form

Public oMat1 As SAPbouiCOM.Matrix

Private oDS_PH_PY310A As SAPbouiCOM.DBDataSource
Private oDS_PH_PY310B As SAPbouiCOM.DBDataSource

Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long

Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY310.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY310_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY310"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
    oForm.DataBrowser.BrowseBy = "DocEntry"
        
    oForm.Freeze True
    Call PH_PY310_CreateItems
    Call PH_PY310_EnableMenus
    Call PH_PY310_SetDocument(oFromDocEntry01)
'    Call PH_PY310_FormResize
    
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

Private Function PH_PY310_CreateItems() As Boolean

    Dim sQry        As String
    Dim i           As Long
    
    Dim oEdit       As SAPbouiCOM.EditText
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns    As SAPbouiCOM.Columns
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY310_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Set oDS_PH_PY310A = oForm.DataSources.DBDataSources("@PH_PY310A")
    Set oDS_PH_PY310B = oForm.DataSources.DBDataSources("@PS_USERDS01")

    '대출일자_S
    Call oForm.DataSources.UserDataSources.Add("LoanDate", dt_DATE)
    Call oForm.Items("LoanDate").Specific.DataBind.SetBound(True, "", "LoanDate")
    '대출일자_E
    
    '대부금액_S
    Call oForm.DataSources.UserDataSources.Add("LoanAmt", dt_SUM, 50)
    Call oForm.Items("LoanAmt").Specific.DataBind.SetBound(True, "", "LoanAmt")
    '대부금액_E

    '총상환금액_S
    Call oForm.DataSources.UserDataSources.Add("TRpmtAmt", dt_SUM, 50)
    Call oForm.Items("TRpmtAmt").Specific.DataBind.SetBound(True, "", "TRpmtAmt")
    '총상환금액_E

    '상환잔액_S
    Call oForm.DataSources.UserDataSources.Add("RmainAmt", dt_SUM, 50)
    Call oForm.Items("RmainAmt").Specific.DataBind.SetBound(True, "", "RmainAmt")
    '상환잔액_E
    
    
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
    
PH_PY310_CreateItems_Error:
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY310_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY310_EnableMenus()
    
    On Error GoTo PH_PY310_EnableMenus_Error
    
    Call oForm.EnableMenu("1283", False)         '// 삭제
    Call oForm.EnableMenu("1287", False)          '// 복제
'    Call oForm.EnableMenu("1286", True)         '// 닫기
    Call oForm.EnableMenu("1284", True)         '// 취소
    Call oForm.EnableMenu("1293", True)         '// 행삭제
    
    Exit Sub
    
PH_PY310_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY310_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY310_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY310_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY310_FormItemEnabled
'        Call PH_PY310_AddMatrixRow
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY310_FormItemEnabled
        oForm.Items("DocEntry").Specific.Value = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY310_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY310_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY310_FormItemEnabled()
    On Error GoTo PH_PY310_FormItemEnabled_Error
    
    Dim oCombo As SAPbouiCOM.ComboBox
    
    Call oForm.Freeze(True)
    If (oForm.Mode = fm_ADD_MODE) Then
        oForm.Items("CLTCOD").Enabled = True
        oForm.Items("CntcCode").Enabled = True
        oForm.Items("LoanDoc").Enabled = True
        oForm.Items("RpmtDate").Enabled = True
        oForm.Items("RpmtAmt").Enabled = True
        oForm.Items("RpmtInt").Enabled = True
        oForm.Items("DocEntry").Enabled = False
        oForm.Items("btnAdd").Enabled = True
        
        '폼 DocEntry 세팅
        Call PH_PY310_FormClear
        
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        
        '대출일자,대출금액,총상환금액,상환잔액 초기화
        oForm.Items("LoanDate").Specific.Value = ""
        oForm.Items("LoanAmt").Specific.Value = ""
        oForm.Items("TRpmtAmt").Specific.Value = ""
        oForm.Items("RmainAmt").Specific.Value = ""
        
        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", False)    '//문서추가
        
    ElseIf (oForm.Mode = fm_FIND_MODE) Then
        oForm.Items("CLTCOD").Enabled = True
        oForm.Items("CntcCode").Enabled = True
        oForm.Items("LoanDoc").Enabled = True
        oForm.Items("RpmtDate").Enabled = True
        oForm.Items("RpmtAmt").Enabled = True
        oForm.Items("RpmtInt").Enabled = True
        oForm.Items("DocEntry").Enabled = True
        oForm.Items("btnAdd").Enabled = True
        
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        
        '대출일자,대출금액,총상환금액,상환잔액 초기화
        oForm.Items("LoanDate").Specific.Value = ""
        oForm.Items("LoanAmt").Specific.Value = ""
        oForm.Items("TRpmtAmt").Specific.Value = ""
        oForm.Items("RmainAmt").Specific.Value = ""
    
        Call oForm.EnableMenu("1281", False)    '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가
        
    ElseIf (oForm.Mode = fm_OK_MODE) Then
        oForm.Items("CLTCOD").Enabled = False
        oForm.Items("CntcCode").Enabled = False
        oForm.Items("LoanDoc").Enabled = False
        oForm.Items("RpmtDate").Enabled = False
        oForm.Items("RpmtAmt").Enabled = False
        oForm.Items("RpmtInt").Enabled = False
        oForm.Items("DocEntry").Enabled = False
        oForm.Items("btnAdd").Enabled = False
         
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD", False)
        
        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가

    End If
    Call oForm.Freeze(False)
    Exit Sub
    
PH_PY310_FormItemEnabled_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY310_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub Raise_FormItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
    On Error GoTo Raise_FormItemEvent_Error
    
    Dim sQry As String
    Dim i As Long
    Dim oCombo As SAPbouiCOM.ComboBox
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    Dim loopCount As Integer 'For Loop 용 (VALIDATE Event에서 사용)
    
    Dim InsertedDocEntry As String '입력된 문서의 문서번호를 리턴

    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Select Case pval.EventType
        Case et_ITEM_PRESSED: '//1
            
            If pval.BeforeAction = True Then
                If pval.ItemUID = "1" Then
                    If oForm.Mode = fm_ADD_MODE Then
                    
                        If PH_PY310_DataValidCheck = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        
                        Call PH_PY310_UpdateRepaymentData
                        
                        '//해야할일 작업
                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
                    
                        If PH_PY310_DataValidCheck = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        '//해야할일 작업
                        
                    ElseIf oForm.Mode = fm_OK_MODE Then
                    End If
                    
                ElseIf pval.ItemUID = "btnAdd" Then
                
                    If oForm.Mode = fm_ADD_MODE Then
                    
                        If PH_PY310_CheckRepaymentData = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If

                        Call PH_PY310_AddRepaymentData
                        
                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
                    
'                        If PH_PY310_DataValidCheck = False Then
'                            BubbleEvent = False
'                        End If
                        
                    ElseIf oForm.Mode = fm_OK_MODE Then
                    
                    
                    End If
                    
                ElseIf pval.ItemUID = "btnTest" Then
                
                    If oForm.Mode = fm_ADD_MODE Then
                    
                        Call PH_PY310_UpdateRepaymentData
                        
                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
                        
                    ElseIf oForm.Mode = fm_OK_MODE Then
                    
                    End If
                    
                End If
            ElseIf pval.BeforeAction = False Then
                If pval.ItemUID = "1" Then
                    If oForm.Mode = fm_ADD_MODE Then
                        If pval.ActionSuccess = True Then
                        
'                            oForm.Mode = fm_FIND_MODE
'
'                            Call PH_PY310_FormItemEnabled
'                            InsertedDocEntry = PH_PY310_GetMaxDocEntry
'                            Call CLTCOD_Select(oForm, "CLTCOD")
'                            Call oDS_PH_PY310A.setValue("DocEntry", 0, InsertedDocEntry)
'                            oForm.Items("1").CLICK ct_Regular
'
''                            Call PH_PY310_FormItemEnabled
''                            Call PH_PY310_AddMatrixRow

                            oForm.Mode = fm_OK_MODE
                            Call Sbo_Application.ActivateMenuItem("1291") '//이동(최종데이타)

                        End If
                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY310_FormItemEnabled
'                            Call PH_PY310_AddMatrixRow
                        End If
                    ElseIf oForm.Mode = fm_OK_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY310_FormItemEnabled
                            Call PH_PY310_GetLoanData(oForm.Items("LoanDoc").Specific.Value) '대출일자, 대부금액, 총상환금액, 상환잔액 표시
                            Call PH_PY310_MTX01
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
            
                ElseIf pval.ItemUID = "LoanDoc" And pval.CharPressed = "9" Then
                
                    If oForm.Items("LoanDoc").Specific.Value = "" Then
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
                        
                            Call oDS_PH_PY310A.setValue("U_CntcName", 0, MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & oForm.Items(pval.ItemUID).Specific.Value & "'"))
                        
                        Case "LoanDoc"
                            
                            If oForm.Items("CntcCode").Specific.Value = "" Then
                            
                                Call oDS_PH_PY310A.setValue("U_CntcCode", 0, MDC_GetData.Get_ReData("U_CntcCode", "DocEntry", "[@PH_PY309A]", "'" & oForm.Items("LoanDoc").Specific.Value & "'"))
                                Call oDS_PH_PY310A.setValue("U_CntcName", 0, MDC_GetData.Get_ReData("U_CntcName", "DocEntry", "[@PH_PY309A]", "'" & oForm.Items("LoanDoc").Specific.Value & "'"))
                            
                            End If
                            
                            Call PH_PY310_MTX01
                            Call PH_PY310_GetLoanData(oForm.Items("LoanDoc").Specific.Value) '대출일자, 대부금액, 총상환금액, 상환잔액 표시
                            
                        Case "Mat01"
                        
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

                Call PH_PY310_FormItemEnabled
                'Call PH_PY310_AddMatrixRow
                Call oMat1.AutoResizeColumns
                
'                Call PH_PY310_GetLoanData(oForm.Items("LoanDoc").Specific.Value) '대출일자, 대부금액, 총상환금액, 상환잔액 표시
'                Call PH_PY310_MTX01
                
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
                Set oDS_PH_PY310A = Nothing
                Set oDS_PH_PY310B = Nothing
                
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
'                    Call MDC_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY310A", "Code")
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
                Call AuthorityCheck(oForm, "CLTCOD", "@PH_PY310A", "DocEntry")      '//접속자 권한에 따른 사업장 보기
        End Select
    ElseIf (pval.BeforeAction = False) Then
        Select Case pval.MenuUID
            Case "1283":
                oForm.Mode = fm_ADD_MODE
                Call PH_PY310_FormItemEnabled
'                Call PH_PY310_AddMatrixRow
            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//문서찾기
                Call PH_PY310_FormItemEnabled
                oForm.Items("DocEntry").CLICK ct_Regular
            Case "1282":    '//문서추가
                Call PH_PY310_FormItemEnabled
'                Call PH_PY310_AddMatrixRow
            Case "1288", "1289", "1290", "1291":
                Call PH_PY310_FormItemEnabled
                Call PH_PY310_GetLoanData(oForm.Items("LoanDoc").Specific.Value) '대출일자, 대부금액, 총상환금액, 상환잔액 표시
                Call PH_PY310_MTX01
            Case "1293" '// 행삭제

'                 If oMat1.RowCount <> oMat1.VisualRowCount Then
'                    oMat1.FlushToDataSource
'
'                    While (i <= oDS_PH_PY310B.Size - 1)
'                        If oDS_PH_PY310B.GetValue("U_LineNum", i) = "" Then
'                            oDS_PH_PY310B.RemoveRecord (i)
'                            i = 0
'                        Else
'                            i = i + 1
'                        End If
'                    Wend
'
'                    For i = 0 To oDS_PH_PY310B.Size
'                        Call oDS_PH_PY310B.setValue("U_LineNum", i, i + 1)
'                    Next i
'
'                    oMat1.LoadFromDataSource
'                End If
                
            '복제
            Case "1287"
            
                Call oForm.Freeze(True)
                Call oDS_PH_PY310A.setValue("DocEntry", 0, "")
                
                For i = 0 To oMat1.VisualRowCount - 1
                    Call oMat1.FlushToDataSource
                    Call oDS_PH_PY310B.setValue("DocEntry", i, "")
                    Call oDS_PH_PY310B.setValue("U_PayYN", i, "N")
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

Sub PH_PY310_AddMatrixRow()
    Dim oRow    As Long
    
    On Error GoTo PH_PY310_AddMatrixRow_Error
    
    Call oForm.Freeze(True)
    
    '//[Mat1]
    oMat1.FlushToDataSource
    oRow = oMat1.VisualRowCount
    
    If oMat1.VisualRowCount > 0 Then
        If Trim(oDS_PH_PY310B.GetValue("U_LineNum", oRow - 1)) <> "" Then
            If oDS_PH_PY310B.Size <= oMat1.VisualRowCount Then
                oDS_PH_PY310B.InsertRecord (oRow)
            End If
            oDS_PH_PY310B.Offset = oRow
            oDS_PH_PY310B.setValue "U_LineNum", oRow, oRow + 1
            oDS_PH_PY310B.setValue "U_ColDt01", oRow, ""
            oDS_PH_PY310B.setValue "U_ColSum01", oRow, 0
            oDS_PH_PY310B.setValue "U_ColSum02", oRow, 0
            oDS_PH_PY310B.setValue "U_ColSum03", oRow, 0
            oDS_PH_PY310B.setValue "U_ColReg01", oRow, "Y"
            oDS_PH_PY310B.setValue "U_ColReg02", oRow, ""
            oMat1.LoadFromDataSource
        Else
            oDS_PH_PY310B.Offset = oRow - 1
            oDS_PH_PY310B.setValue "U_LineNum", oRow - 1, oRow
            oDS_PH_PY310B.setValue "U_ColDt01", oRow - 1, ""
            oDS_PH_PY310B.setValue "U_ColSum01", oRow - 1, 0
            oDS_PH_PY310B.setValue "U_ColSum02", oRow - 1, 0
            oDS_PH_PY310B.setValue "U_ColSum03", oRow - 1, 0
            oDS_PH_PY310B.setValue "U_ColReg01", oRow - 1, "Y"
            oDS_PH_PY310B.setValue "U_ColReg02", oRow - 1, ""
            oMat1.LoadFromDataSource
        End If
    ElseIf oMat1.VisualRowCount = 0 Then
        oDS_PH_PY310B.Offset = oRow
        oDS_PH_PY310B.setValue "U_LineNum", oRow, oRow + 1
        oDS_PH_PY310B.setValue "U_ColDt01", oRow, 0
        oDS_PH_PY310B.setValue "U_ColSum01", oRow, 0
        oDS_PH_PY310B.setValue "U_ColSum02", oRow, 0
        oDS_PH_PY310B.setValue "U_ColSum03", oRow, 0
        oDS_PH_PY310B.setValue "U_ColReg01", oRow, "Y"
        oDS_PH_PY310B.setValue "U_ColReg02", oRow, ""
        oMat1.LoadFromDataSource
    End If
    
    Call oForm.Freeze(False)
    Exit Sub
PH_PY310_AddMatrixRow_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY310_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY310_FormClear()
On Error GoTo PH_PY310_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY310'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.Value = 1
    Else
        oForm.Items("DocEntry").Specific.Value = DocEntry
        
        '방금 입력한 문서를 조회하기 위해 버튼 클릭
'        oForm.Items("1").CLICK ct_Regular
        
    End If
    Exit Sub
PH_PY310_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY310_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY310_DataValidCheck() As Boolean
On Error GoTo PH_PY310_DataValidCheck_Error
    PH_PY310_DataValidCheck = False
    Dim i       As Long
    Dim sQry    As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    '사업장
    If Trim(oDS_PH_PY310A.GetValue("U_CLTCOD", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "사업장은 필수입니다.", bmt_Short, True
        oForm.Items("CLTCOD").CLICK ct_Regular
        PH_PY310_DataValidCheck = False
        Exit Function
    End If
    
    '사번
    If Trim(oDS_PH_PY310A.GetValue("U_CntcCode", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "사번은 필수입니다.", bmt_Short, True
        oForm.Items("CntcCode").CLICK ct_Regular
        PH_PY310_DataValidCheck = False
        Exit Function
    End If
    
    '대부금문서번호
    If Trim(oDS_PH_PY310A.GetValue("U_LoanDoc", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "대부금문서는 필수입니다.", bmt_Short, True
        oForm.Items("LoanDoc").CLICK ct_Regular
        PH_PY310_DataValidCheck = False
        Exit Function
    End If
    
    '상환일자
    If Trim(oDS_PH_PY310A.GetValue("U_RpmtDate", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "상환일자는 필수입니다.", bmt_Short, True
        oForm.Items("RpmtDate").CLICK ct_Regular
        PH_PY310_DataValidCheck = False
        Exit Function
    End If
    
    '상환금액
    If Trim(oDS_PH_PY310A.GetValue("U_RpmtAmt", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "상환금액은 필수입니다.", bmt_Short, True
        oForm.Items("RpmtAmt").CLICK ct_Regular
        PH_PY310_DataValidCheck = False
        Exit Function
    End If
    
    '상환이자
    If Trim(oDS_PH_PY310A.GetValue("U_RpmtInt", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "상환이자는 필수입니다.", bmt_Short, True
        oForm.Items("RpmtInt").CLICK ct_Regular
        PH_PY310_DataValidCheck = False
        Exit Function
    End If
    
    '라인
    If oMat1.VisualRowCount > 0 Then
'        For i = 1 To oMat1.VisualRowCount - 1
'
'            '학교
'            If oMat1.Columns("SchCls").Cells(i).Specific.Value = "" Then
'                Sbo_Application.SetStatusBarMessage "학교는 필수입니다.", bmt_Short, True
'                oMat1.Columns("SchCls").Cells(i).CLICK ct_Regular
'                PH_PY310_DataValidCheck = False
'                Exit Function
'            End If
'
'            '학교명
'            If oMat1.Columns("SchName").Cells(i).Specific.Value = "" Then
'                Sbo_Application.SetStatusBarMessage "학교명은 필수입니다.", bmt_Short, True
'                oMat1.Columns("SchName").Cells(i).CLICK ct_Regular
'                PH_PY310_DataValidCheck = False
'                Exit Function
'            End If
'
'            '학년
'            If oMat1.Columns("Grade").Cells(i).Specific.Value = "" Then
'                Sbo_Application.SetStatusBarMessage "학년은 필수입니다.", bmt_Short, True
'                oMat1.Columns("Grade").Cells(i).CLICK ct_Regular
'                PH_PY310_DataValidCheck = False
'                Exit Function
'            End If
'
'            '회차
'            If oMat1.Columns("Count").Cells(i).Specific.Value = "" Then
'                Sbo_Application.SetStatusBarMessage "회차는 필수입니다.", bmt_Short, True
'                oMat1.Columns("Count").Cells(i).CLICK ct_Regular
'                PH_PY310_DataValidCheck = False
'                Exit Function
'            End If
'
'        Next
    Else
        Sbo_Application.SetStatusBarMessage "라인 데이터가 없습니다.", bmt_Short, True
        PH_PY310_DataValidCheck = False
        Exit Function
    End If
    
'    oMat1.FlushToDataSource
'    '// Matrix 마지막 행 삭제(DB 저장시)
'    If oDS_PH_PY310B.Size > 1 Then oDS_PH_PY310B.RemoveRecord (oDS_PH_PY310B.Size - 1)
'
'    oMat1.LoadFromDataSource
    
    PH_PY310_DataValidCheck = True
    
    Exit Function
    
    Set oRecordSet = Nothing
    
    
PH_PY310_DataValidCheck_Error:
    Set oRecordSet = Nothing
    PH_PY310_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY310_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY310_MTX01()
On Error GoTo PH_PY310_MTX01_Error
    '//메트릭스에 데이터 로드
    
    Dim i As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Dim Param01 As String
    Dim Param02 As String
    Dim Param03 As String
    Dim Param04 As String
    
    Param01 = oForm.Items("LoanDoc").Specific.Value
'    Param02 = oForm.Items("LoanDate").Specific.Value
'    Param03 = oForm.Items("RpmtPrd").Specific.Value
'    Param04 = oForm.Items("Param01").Specific.Value
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet.RecordCount, False)

    Call oForm.Freeze(True)
    
    sQry = "EXEC PH_PY310_01 '" & Param01 & "'"
    Call oRecordSet.DoQuery(sQry)
    
    oMat1.Clear
    oMat1.FlushToDataSource
    oMat1.LoadFromDataSource
    
    If (oRecordSet.RecordCount = 0) Then
        Call oMat1.Clear
        GoTo PH_PY310_MTX01_Exit
    End If

    For i = 0 To oRecordSet.RecordCount - 1
        If i <> 0 Then
            Call oDS_PH_PY310B.InsertRecord(i)
        End If
        
'        If i = oRecordSet.RecordCount Then '마지막 빈행 추가
'
'            oDS_PH_PY310B.Offset = i
'            Call oDS_PH_PY310B.setValue("U_LineNum", i, i + 1) '라인번호
'            Call oDS_PH_PY310B.setValue("U_Cnt", i, "")
'            Call oDS_PH_PY310B.setValue("U_PayDate", i, "")
'            Call oDS_PH_PY310B.setValue("U_RpmtAmt", i, 0)
'            Call oDS_PH_PY310B.setValue("U_TotRpmt", i, 0)
'
'        Else
        
            oDS_PH_PY310B.Offset = i
            Call oDS_PH_PY310B.setValue("U_LineNum", i, i + 1) '라인번호
            Call oDS_PH_PY310B.setValue("U_ColDt01", i, oRecordSet.Fields("RpmtDate").Value)
            Call oDS_PH_PY310B.setValue("U_ColSum01", i, oRecordSet.Fields("RpmtAmt").Value)
            Call oDS_PH_PY310B.setValue("U_ColSum02", i, oRecordSet.Fields("RpmtInt").Value)
            Call oDS_PH_PY310B.setValue("U_ColSum03", i, oRecordSet.Fields("RmainAmt").Value)
            Call oDS_PH_PY310B.setValue("U_ColReg01", i, oRecordSet.Fields("AddYN").Value)
            Call oDS_PH_PY310B.setValue("U_ColReg02", i, oRecordSet.Fields("Cnt").Value)
            
            oRecordSet.MoveNext
        
'        End If
        
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
PH_PY310_MTX01_Exit:
    Set oRecordSet = Nothing
    Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
    Call oForm.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Exit Sub
PH_PY310_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY310_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY310_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY310_Validate_Error
    PH_PY310_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY310A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.Value & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", bmt_Short, True
        PH_PY310_Validate = False
        GoTo PH_PY310_Validate_Exit
    End If
'
    If ValidateType = "수정" Then

    ElseIf ValidateType = "행삭제" Then

    ElseIf ValidateType = "취소" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY310_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY310_Validate_Error:
    PH_PY310_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY310_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY310_GetLoanData(pLoanDoc As Integer)
'******************************************************************************
'Function ID : PH_PY310_GetLoanData()
'해당모듈 : PH_PY310
'기능 : 대출일자, 대출금액, 총상환금액, 상환잔액 조회
'인수 : pLoanDoc:대부금등록 문서번호
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PH_PY310_GetLoanData_Error

    Dim loopCount As Integer
    Dim sQry As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    sQry = "EXEC PH_PY310_02 '" & pLoanDoc & "'"
    
    Call oRecordSet.DoQuery(sQry)
    
    oForm.Items("LoanDate").Specific.Value = oRecordSet.Fields("LoanDate").Value
    oForm.Items("LoanAmt").Specific.Value = oRecordSet.Fields("Loanamt").Value
    oForm.Items("TRpmtAmt").Specific.Value = oRecordSet.Fields("TRpmtAmt").Value
    oForm.Items("RmainAmt").Specific.Value = oRecordSet.Fields("RmainAmt").Value

    Exit Sub
PH_PY310_GetLoanData_Error:
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "PH_PY310_GetLoanData_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PH_PY310_CheckRepaymentData() As Boolean
'******************************************************************************
'Function ID : PH_PY310_CheckRepaymentData()
'해당모듈 : PH_PY310
'기능 : 상환일자, 상환금액, 상환이자 입력 여부 체크
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PH_PY310_CheckRepaymentData_Error

    PH_PY310_CheckRepaymentData = False

    '상환일자
    If Trim(oDS_PH_PY310A.GetValue("U_RpmtDate", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "상환일자는 필수입니다.", bmt_Short, True
        oForm.Items("RpmtDate").CLICK ct_Regular
        PH_PY310_CheckRepaymentData = False
        Exit Function
    End If
    
    '상환금액
    If Trim(oDS_PH_PY310A.GetValue("U_RpmtAmt", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "상환금액은 필수입니다.", bmt_Short, True
        oForm.Items("RpmtAmt").CLICK ct_Regular
        PH_PY310_CheckRepaymentData = False
        Exit Function
    End If
    
    '상환이자
    If Trim(oDS_PH_PY310A.GetValue("U_RpmtInt", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "상환이자는 필수입니다.", bmt_Short, True
        oForm.Items("RpmtInt").CLICK ct_Regular
        PH_PY310_CheckRepaymentData = False
        Exit Function
    End If
    
    PH_PY310_CheckRepaymentData = True

    Exit Function
    
PH_PY310_CheckRepaymentData_Error:
    PH_PY310_CheckRepaymentData = False
    Sbo_Application.SetStatusBarMessage "PH_PY310_PH_PY310_CheckRepaymentData_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Private Sub PH_PY310_AddRepaymentData()
'******************************************************************************
'Function ID : PH_PY310_AddRepaymentData()
'해당모듈 : PH_PY310
'기능 : 개별상환내역을 Matrix에 행 추가
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PH_PY310_AddRepaymentData_Error

    Dim oRow As Long
    
    Call oForm.Freeze(True)
    
    '//[Mat1]
    oMat1.FlushToDataSource
    oRow = oMat1.VisualRowCount
    
    If oMat1.VisualRowCount > 0 Then
        If Trim(oDS_PH_PY310B.GetValue("U_LineNum", oRow - 1)) <> "" Then
            If oDS_PH_PY310B.Size <= oMat1.VisualRowCount Then
                oDS_PH_PY310B.InsertRecord (oRow)
            End If
            oDS_PH_PY310B.Offset = oRow
            Call oDS_PH_PY310B.setValue("U_LineNum", oRow, oRow + 1)
            Call oDS_PH_PY310B.setValue("U_ColDt01", oRow, Trim(oDS_PH_PY310A.GetValue("U_RpmtDate", 0)))
            Call oDS_PH_PY310B.setValue("U_ColSum01", oRow, Trim(oDS_PH_PY310A.GetValue("U_RpmtAmt", 0)))
            Call oDS_PH_PY310B.setValue("U_ColSum02", oRow, Trim(oDS_PH_PY310A.GetValue("U_RpmtInt", 0)))
            Call oDS_PH_PY310B.setValue("U_ColSum03", oRow, (oForm.Items("LoanAmt").Specific.Value - oForm.Items("TRpmtAmt").Specific.Value) - Trim(oDS_PH_PY310A.GetValue("U_RpmtAmt", 0)))
            Call oDS_PH_PY310B.setValue("U_ColReg01", oRow, "Y")
            Call oDS_PH_PY310B.setValue("U_ColReg02", oRow, oDS_PH_PY310B.GetValue("U_ColReg02", oRow - 1) + 1) '이전 행의 회차 + 1
            Call oMat1.LoadFromDataSource
        Else
            oDS_PH_PY310B.Offset = oRow - 1
            Call oDS_PH_PY310B.setValue("U_LineNum", oRow - 1, oRow)
            Call oDS_PH_PY310B.setValue("U_ColDt01", oRow - 1, Trim(oDS_PH_PY310A.GetValue("U_RpmtDate", 0)))
            Call oDS_PH_PY310B.setValue("U_ColSum01", oRow - 1, Trim(oDS_PH_PY310A.GetValue("U_RpmtAmt", 0)))
            Call oDS_PH_PY310B.setValue("U_ColSum02", oRow - 1, Trim(oDS_PH_PY310A.GetValue("U_RpmtInt", 0)))
            Call oDS_PH_PY310B.setValue("U_ColSum03", oRow - 1, (oForm.Items("LoanAmt").Specific.Value - oForm.Items("TRpmtAmt").Specific.Value) - Trim(oDS_PH_PY310A.GetValue("U_RpmtAmt", 0)))
            Call oDS_PH_PY310B.setValue("U_ColReg01", oRow - 1, "Y")
            Call oDS_PH_PY310B.setValue("U_ColReg02", oRow - 1, oDS_PH_PY310B.GetValue("U_ColReg02", oRow - 2) + 1) '이전 행의 회차 + 1
            Call oMat1.LoadFromDataSource
        End If
    ElseIf oMat1.VisualRowCount = 0 Then
        oDS_PH_PY310B.Offset = oRow
        Call oDS_PH_PY310B.setValue("U_LineNum", oRow, oRow + 1)
        Call oDS_PH_PY310B.setValue("U_ColDt01", oRow, Trim(oDS_PH_PY310A.GetValue("U_RpmtDate", 0)))
        Call oDS_PH_PY310B.setValue("U_ColSum01", oRow, Trim(oDS_PH_PY310A.GetValue("U_RpmtAmt", 0)))
        Call oDS_PH_PY310B.setValue("U_ColSum02", oRow, Trim(oDS_PH_PY310A.GetValue("U_RpmtInt", 0)))
        Call oDS_PH_PY310B.setValue("U_ColSum03", oRow, (oForm.Items("LoanAmt").Specific.Value - oForm.Items("TRpmtAmt").Specific.Value) - Trim(oDS_PH_PY310A.GetValue("U_RpmtAmt", 0)))
        Call oDS_PH_PY310B.setValue("U_ColReg01", oRow, "Y")
        Call oDS_PH_PY310B.setValue("U_ColReg02", oRow, "1") '첫 행은 무조건 1회차
        Call oMat1.LoadFromDataSource
    End If
    
    Call oForm.Freeze(False)
    Exit Sub
PH_PY310_AddRepaymentData_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY310_AddRepaymentData_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY310_UpdateRepaymentData()
'******************************************************************************
'Function ID : PH_PY310_UpdateRepaymentData()
'해당모듈 : PH_PY310
'기능 : 대부금개별상환 내역을 Z_PH_PY310에 INSERT, @PH_PY309B에 UPDATE
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PH_PY310_UpdateRepaymentData_Error

    Dim loopCount As Integer
    Dim sQry As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Dim CLTCOD As String '사업장
    Dim CntcCode As String '사번
    Dim LoanDoc As Integer '대부금문서번호
    Dim RpmtDate As String '상환일자
    Dim RpmtAmt As Double '상환금액
    Dim RpmtInt As Double '상환이자
    Dim RmainAmt As Double '상환잔액
    Dim Cnt As Integer '회차
    
    For loopCount = 0 To oMat1.VisualRowCount - 1
    
        If Trim(oDS_PH_PY310B.GetValue("U_ColReg01", loopCount)) = "Y" Then
        
            CLTCOD = oDS_PH_PY310A.GetValue("U_CLTCOD", 0)
            CntcCode = oDS_PH_PY310A.GetValue("U_CntcCode", 0)
            LoanDoc = oDS_PH_PY310A.GetValue("U_LoanDoc", 0)
            RpmtDate = oDS_PH_PY310B.GetValue("U_ColDt01", loopCount)
            RpmtAmt = oDS_PH_PY310B.GetValue("U_ColSum01", loopCount)
            RpmtInt = oDS_PH_PY310B.GetValue("U_ColSum02", loopCount)
            RmainAmt = oDS_PH_PY310B.GetValue("U_ColSum03", loopCount)
            Cnt = oDS_PH_PY310B.GetValue("U_ColReg02", loopCount)
            
            sQry = "EXEC PH_PY310_03 '" & CLTCOD & "','" & CntcCode & "','" & LoanDoc & "','" & RpmtDate & "','" & RpmtAmt & "','" & RpmtInt & "','" & RmainAmt & "','" & Cnt & "'"
    
            Call oRecordSet.DoQuery(sQry)
        
        End If
    
    Next

    Set oRecordSet = Nothing
    Exit Sub
PH_PY310_UpdateRepaymentData_Error:
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "PH_PY310_UpdateRepaymentData_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PH_PY310_GetMaxDocEntry() As String
'******************************************************************************
'Function ID : PH_PY310_GetMaxDocEntry()
'해당모듈 : PH_PY310
'기능 : 가장 최근 등록된 문서번호를 조회
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PH_PY310_GetMaxDocEntry_Error

    Dim loopCount As Integer
    Dim sQry As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    sQry = "SELECT MAX(AutoKey) - 1 AS [MaxDocEntry] FROM ONNM WHERE ObjectCode = 'PH_PY310'"
    
    Call oRecordSet.DoQuery(sQry)
    
    PH_PY310_GetMaxDocEntry = oRecordSet.Fields("MaxDocEntry").Value
    
    Set oRecordSet = Nothing
    Exit Function
PH_PY310_GetMaxDocEntry_Error:
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "PH_PY310_GetMaxDocEntry_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function
