VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY118"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY118.cls
'//  Module         : 급여관리 >
'//  Desc           : 급상여 Email 발송
'//********************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form

''// 매트릭스 사용시
Public oMat1 As SAPbouiCOM.Matrix

Private oDS_PH_PY118B As SAPbouiCOM.DBDataSource

Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long

Private oJOBYMM         As String
Private oJOBTYP         As String
Private oJOBGBN         As String
Private oPAYSEL         As String
Private oCLTCOD         As String
Private oMSTBRK         As String
Private oMSTDPT         As String
Private oMSTCOD         As String

Private sHtml           As Variant
Private ArrPayHead(36)  As String
Private ArrPayAmt(36)   As Double
Private ArrSubHead(36)  As String
Private ArrSubAmt(36)   As Double
Private ArrGntHead(18)  As String
Private ArrGntAmt(18)   As Double
Private sMSTCOD         As String
Private sMSTNAM         As String
Private sDPTNAM         As String
Private sPOSITION       As String
Private sTOTPAY         As Double
Private sTOTGON         As Double
Private sSILJIG         As Double
Private sTOEmail        As String
Private sFrEmail        As String
Private sFrSMTP         As String
Private sFrSMTPSrv      As String
Private sFrSMTPPort     As String
Private sFrPWD          As String
Private oPRTTIL         As String
Private oPRTSUB         As String
Private oCLTNAM         As String
Private oREMARK         As String
Private oPrtChk         As Boolean

Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY118.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY118_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY118"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_FIND_MODE
'    oForm.DataBrowser.BrowseBy = "Code"
        
    oForm.Freeze True
    Call PH_PY118_CreateItems
    Call PH_PY118_EnableMenus
    Call PH_PY118_SetDocument(oFromDocEntry01)
'    Call PH_PY118_FormResize
    
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

Private Function PH_PY118_CreateItems() As Boolean

    Dim sQry        As String
    Dim i           As Long
    
    Dim oCheck      As SAPbouiCOM.CheckBox
    Dim oEdit       As SAPbouiCOM.EditText
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns    As SAPbouiCOM.Columns
    Dim optBtn      As SAPbouiCOM.OptionBtn
    Dim oLink           As SAPbouiCOM.LinkedButton
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY118_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    
    '//----------------------------------------------------------------------------------------------
    '// 데이터셋정의
    '//----------------------------------------------------------------------------------------------
'    '//테이블이 있을경우 데이터셋(Matrix)
    Set oDS_PH_PY118B = oForm.DataSources.DBDataSources("@PH_PY118B")


    '// 귀속년월
    Call oForm.DataSources.UserDataSources.Add("JOBYMM", dt_SHORT_TEXT, 6)
    Set oEdit = oForm.Items("JOBYMM").Specific
    oEdit.DataBind.SetBound True, "", "JOBYMM"
    oForm.DataSources.UserDataSources.Item("JOBYMM").ValueEx = Format(Date, "YYYYMM")
    
    '//지급종류
    Call oForm.DataSources.UserDataSources.Add("JOBTYP", dt_SHORT_TEXT)
    Set oCombo = oForm.Items("JOBTYP").Specific
    oCombo.DataBind.SetBound True, "", "JOBTYP"
    
    oCombo.ValidValues.Add "1", "급여"
    oCombo.ValidValues.Add "2", "상여"
    oCombo.Select "0", psk_Index
    oForm.Items("JOBTYP").DisplayDesc = True
    
    '//지급구분
    Call oForm.DataSources.UserDataSources.Add("JOBGBN", dt_SHORT_TEXT)
    Set oCombo = oForm.Items("JOBGBN").Specific
    oCombo.DataBind.SetBound True, "", "JOBGBN"
    
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P212' AND U_UseYN= 'Y'"
    Call SetReDataCombo(oForm, sQry, oCombo)
    oCombo.Select "0", psk_Index
    oForm.Items("JOBGBN").DisplayDesc = True
    
    
    '//지급대상자구분
    Call oForm.DataSources.UserDataSources.Add("PAYSEL", dt_SHORT_TEXT)
    Set oCombo = oForm.Items("PAYSEL").Specific
    oCombo.DataBind.SetBound True, "", "PAYSEL"
    
    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P213' ORDER BY CAST(U_Code AS NUMERIC) "
    Call SetReDataCombo(oForm, sQry, oCombo)
    oCombo.ValidValues.Add "%", "전체"
    oCombo.Select "%", psk_ByValue
    
    oForm.Items("PAYSEL").DisplayDesc = True
    
    '//부서
    Call oForm.DataSources.UserDataSources.Add("MSTDPT", dt_SHORT_TEXT)
    Set oCombo = oForm.Items("MSTDPT").Specific
    oCombo.DataBind.SetBound True, "", "MSTDPT"

    oForm.Items("MSTDPT").DisplayDesc = True
    
    '//사업장
    Call oForm.DataSources.UserDataSources.Add("CLTCOD", dt_SHORT_TEXT, 6)
    Set oCombo = oForm.Items("CLTCOD").Specific
    oCombo.DataBind.SetBound True, "", "CLTCOD"
    
    Call CLTCOD_Select(oForm, "CLTCOD")
    oForm.Items("CLTCOD").DisplayDesc = True
    
    '//사원코드
    Call oForm.DataSources.UserDataSources.Add("MSTCOD", dt_SHORT_TEXT, 8)
    Set oEdit = oForm.Items("MSTCOD").Specific
    oEdit.DataBind.SetBound True, "", "MSTCOD"
    
    '//사원명
    Call oForm.DataSources.UserDataSources.Add("MSTNAM", dt_SHORT_TEXT, 30)
    Set oEdit = oForm.Items("MSTNAM").Specific
    oEdit.DataBind.SetBound True, "", "MSTNAM"
    
    '//보내는사람주소
    Call oForm.DataSources.UserDataSources.Add("FrEmail", dt_SHORT_TEXT, 50)
    Set oEdit = oForm.Items("FrEmail").Specific
    oEdit.DataBind.SetBound True, "", "FrEmail"
    
    '//SMTP server
    Call oForm.DataSources.UserDataSources.Add("FrSMTP", dt_SHORT_TEXT, 50)
    Set oEdit = oForm.Items("FrSMTP").Specific
    oEdit.DataBind.SetBound True, "", "FrSMTP"
    
    '//패스워드
    Call oForm.DataSources.UserDataSources.Add("FrPWD", dt_SHORT_TEXT, 30)
    Set oEdit = oForm.Items("FrPWD").Specific
    oEdit.DataBind.SetBound True, "", "FrPWD"
    
    '//공지사항
    Call oForm.DataSources.UserDataSources.Add("Remark", dt_LONG_TEXT)
    Set oEdit = oForm.Items("Remark").Specific
    oEdit.DataBind.SetBound True, "", "Remark"
    
    '//익명인증사용
    Call oForm.DataSources.UserDataSources.Add("AUTCHK", dt_SHORT_TEXT, 1)
    Set oCheck = oForm.Items("AUTCHK").Specific
    oCheck.DataBind.SetBound True, "", "AUTCHK"
    

    
'/ 6.기본설정 가져오기
    sQry = " SELECT U_FrSMTP, U_FrEMAIL, U_FrPWD FROM [@PH_PY118A] WHERE Code= '1'"
    oRecordSet.DoQuery sQry
    Do Until oRecordSet.EOF
        oForm.DataSources.UserDataSources("FrSMTP").VALUE = oRecordSet.Fields(0).VALUE
        oForm.DataSources.UserDataSources("FrEMAIL").VALUE = oRecordSet.Fields(1).VALUE
        oForm.DataSources.UserDataSources("FrPWD").VALUE = oRecordSet.Fields(2).VALUE
        oRecordSet.MoveNext
    Loop
    
'/ Matrix
    Set oMat1 = oForm.Items("Mat1").Specific
    MDC_SetMod.PAY_Matrix_AddCol oMat1, "Col08", 121, "포함", 50, True, False, True, "@PH_PY118B", "U_Col08"
    MDC_SetMod.PAY_Matrix_AddCol oMat1, "Col01", 16, "부서", 80, True, False, True, "@PH_PY118B", "U_Col01"
    MDC_SetMod.PAY_Matrix_AddCol oMat1, "Col02", 16, "직책", 80, True, False, True, "@PH_PY118B", "U_Col02"
    MDC_SetMod.PAY_Matrix_AddCol oMat1, "Col03", 16, "입사일", 70, True, False, True, "@PH_PY118B", "U_Col03"
    MDC_SetMod.PAY_Matrix_AddCol oMat1, "Col04", 16, "퇴사일", 70, True, False, True, "@PH_PY118B", "U_Col04"
    MDC_SetMod.PAY_Matrix_AddCol oMat1, "Col05", 16, "실지급액", 80, True, True, True, "@PH_PY118B", "U_Col05"
    MDC_SetMod.PAY_Matrix_AddCol oMat1, "Col06", 16, "Email주소", 80, True, False, True, "@PH_PY118B", "U_Col06"
    MDC_SetMod.PAY_Matrix_AddCol oMat1, "Col07", 16, "확인", 50, True, False, True, "@PH_PY118B", "U_Col07"
    
'    Set oColumn = oMat1.Columns("EMPID")
'    Set oLink = oColumn.ExtendedObject
'    oLink.LinkedObject = lf_Employee
    
    '/ Check 버튼
    Set oCheck = oForm.Items("AUTCHK").Specific
    oCheck.ValOff = "N": oCheck.ValOn = "Y"
    
    oMat1.AutoResizeColumns
    
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
    
PH_PY118_CreateItems_Error:
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY118_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY118_EnableMenus()
    
    On Error GoTo PH_PY118_EnableMenus_Error

    Call oForm.EnableMenu("1282", False)     '//추가
    Call oForm.EnableMenu("1283", False)     '//제거
    Call oForm.EnableMenu("1284", False)     '//취소
    Call oForm.EnableMenu("1293", False)     '//행삭제
    
    Exit Sub
    
PH_PY118_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY118_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY118_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY118_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY118_FormItemEnabled
        Call PH_PY118_AddMatrixRow
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY118_FormItemEnabled
        oForm.Items("Code").Specific.VALUE = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY118_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY118_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY118_FormItemEnabled()
    On Error GoTo PH_PY118_FormItemEnabled_Error
    
    
    Call oForm.Freeze(True)
    If oForm.Items("Btn2").Visible = True Then
        oForm.Items("FrSMTP").Visible = True
        oForm.Items("FrEmail").Visible = True
        oForm.Items("FrPWD").Visible = True
        oForm.Items("s06").Visible = True
        oForm.Items("s07").Visible = True
        oForm.Items("s08").Visible = True
        oForm.Items("Btn2").Visible = False
        oForm.Items("Btn3").Visible = True
    Else
        oForm.Items("Remark").CLICK ct_Regular
        oForm.Items("FrSMTP").Visible = False
        oForm.Items("FrEmail").Visible = False
        oForm.Items("FrPWD").Visible = False
        oForm.Items("s06").Visible = False
        oForm.Items("s07").Visible = False
        oForm.Items("s08").Visible = False
        oForm.Items("Btn2").Visible = True
        oForm.Items("Btn3").Visible = False
    End If
    Call oForm.Freeze(False)
    Exit Sub
    
PH_PY118_FormItemEnabled_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY118_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
                    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
                    '//추가및 업데이트시에
                    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
                    If oForm.Mode = fm_FIND_MODE Then
                        'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
                        '//찾기모드일때 조건조회
                        'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
                        If Execution_Process = False Then
                            BubbleEvent = False
                            Exit Sub
                        Else
                            BubbleEvent = False
                            oMat1.AutoResizeColumns
                            Exit Sub
                        End If
                    End If
                ElseIf pval.ItemUID = "CBtn1" Then
                    If oForm.Items("MSTCOD").Enabled = True Then
                        oForm.Items("MSTCOD").CLICK ct_Regular
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                End If
            ElseIf pval.BeforeAction = False Then
                If pval.ItemUID = "Btn1" Then
                    Call eMail_Process
                ElseIf pval.ItemUID = "Btn2" Or pval.ItemUID = "Btn3" Then
                    Call PH_PY118_FormItemEnabled
                    If pval.ItemUID = "Btn3" Then  '/ 임시테이블에 저장하기
                        Call SetSendMail
                    End If
                End If
            End If
     
'----------------------------------------------------------
        Case et_KEY_DOWN: '//2
'            'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
'            '//추가모드에서 코드이벤트가 코드에서 일어 났을때
'            'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
'             If pval.BeforeAction = True And pval.ColUid = "STRDAT" And pval.CharPressed = 9 Then
'                If Trim$(oForm.Items(pval.ItemUID).Specific.Value) = "" Then
'                    Call SBO_Application.StatusBar.SetText("조회시작일을 입력하여 주십시오.", bmt_Short, smt_Error)
'                    BubbleEvent = False
'                End If
'             ElseIf pval.BeforeAction = True And pval.ColUid = "ENDDAT" And pval.CharPressed = 9 Then
'                If Trim$(oForm.Items(pval.ItemUID).Specific.Value) = "" Then
'                    Call SBO_Application.StatusBar.SetText("조회종료일을 입력하여 주십시오.", bmt_Short, smt_Error)
'                    BubbleEvent = False
'                End If
'             End If
'----------------------------------------------------------
        Case et_GOT_FOCUS: '//3
'            Select Case pval.ItemUid
'            Case "Mat1", "Grid1"
'                If pval.Row > 0 Then
'                    oLastItemUID = pval.ItemUid
'                    oLastColUID = pval.ColUID
'                    oLastColRow = pval.Row
'                End If
'            Case Else
'                oLastItemUID = pval.ItemUid
'                oLastColUID = ""
'                oLastColRow = 0
'            End Select
'----------------------------------------------------------
        Case et_LOST_FOCUS: '//4

'----------------------------------------------------------
        Case et_COMBO_SELECT: '//5
            Call oForm.Freeze(True)
            If pval.BeforeAction = True Then

            ElseIf pval.BeforeAction = False Then
                If pval.ItemChanged = True Then
                    If pval.ItemUID = "CLTCOD" Then
                        '//기본사항 - 부서 (사업장에 따른 부서변경)
                        Set oCombo = oForm.Items("MSTDPT").Specific

                        sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
                        sQry = sQry & " WHERE Code = '1' AND U_Char2 = '" & Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) & "'"
                        sQry = sQry & " ORDER BY U_Code"
                        
                        Call SetReDataCombo(oForm, sQry, oCombo)
                        
                        If oCombo.ValidValues.Count > 0 Then
                            oCombo.ValidValues.Add "%", "전체"
                            oCombo.Select "%", psk_ByValue
                        End If
                        
                        oForm.Items("MSTDPT").DisplayDesc = True
                    End If
                End If
            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_CLICK: '//6
            If pval.FormUID = oForm.uniqueID And pval.BeforeAction = False And pval.ItemUID = "Mat1" _
                And pval.ColUID = "Col11" Then
                If pval.Row = 0 Then
                    Call PrintChk
                End If
            End If
'----------------------------------------------------------
        Case et_DOUBLE_CLICK: '//7
             If pval.FormUID = oForm.uniqueID And pval.BeforeAction = False And pval.ItemUID = "Mat1" _
                And pval.ColUID = "Col11" Then
                If pval.Row = 0 Then
                    Call PrintChk
                End If
            End If
'----------------------------------------------------------
        Case et_MATRIX_LINK_PRESSED '//8

'----------------------------------------------------------
        Case et_MATRIX_COLLAPSE_PRESSED '//9
'----------------------------------------------------------
        Case et_VALIDATE: '//10
             If pval.BeforeAction = False And pval.ItemChanged = True Then
                Select Case pval.ItemUID
                Case "MSTCOD"
                    If oForm.Items("MSTCOD").Specific.String = "" Then
                        oForm.Items("MSTNAM").Specific.VALUE = ""
                    Else
                        oForm.Items("MSTNAM").Specific.VALUE = MDC_SetMod.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & Trim$(oForm.Items("MSTCOD").Specific.String) & "'", "")
                    End If
                    oForm.Items("MSTNAM").Update
                
                Case "STRDAT"
                    If oForm.Items("STRDAT").Specific.String = "" Then
                        If oForm.Items("STRDAT").Specific.String = "" Then
                            oForm.Items("STRDAT").Specific.VALUE = Format$(Now, "YYYYMMDD")
                        End If
                    End If
                
                Case "ENDDAT"
                    If oForm.Items("ENDDAT").Specific.String = "" Then
                        If oForm.Items("ENDDAT").Specific.String = "" Then
                            oForm.Items("ENDDAT").Specific.VALUE = Format$(Now, "YYYYMMDD")
                        End If
                    End If
                End Select
            End If
'----------------------------------------------------------
        Case et_MATRIX_LOAD: '//11
'            If pval.BeforeAction = True Then
'            ElseIf pval.BeforeAction = False Then
'                oMat1.LoadFromDataSource
'
'                Call PH_PY118_FormItemEnabled
'                Call PH_PY118_AddMatrixRow
'
'            End If
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
                Set oDS_PH_PY118B = Nothing

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
'            If pval.BeforeAction = True Then
'
'            ElseIf pval.Before_Action = False Then
'
'            End If
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
                Call PH_PY118_FormItemEnabled
                Call PH_PY118_AddMatrixRow
            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//문서찾기
                Call PH_PY118_FormItemEnabled
                Call PH_PY118_AddMatrixRow
            Case "1282":    '//문서추가
                Call PH_PY118_FormItemEnabled
                Call PH_PY118_AddMatrixRow
            Case "1288", "1289", "1290", "1291":
                Call PH_PY118_FormItemEnabled
            Case "1293" '// 행삭제
'                '// [MAT1 용]
'                 If oMat1.RowCount <> oMat1.VisualRowCount Then
'                    oMat1.FlushToDataSource
'
'                    While (i <= oDS_PH_PY118B.Size - 1)
'                        If oDS_PH_PY118B.GetValue("U_FILD01", i) = "" Then
'                            oDS_PH_PY118B.RemoveRecord (i)
'                            i = 0
'                        Else
'                            i = i + 1
'                        End If
'                    Wend
'
'                    For i = 0 To oDS_PH_PY118B.Size
'                        Call oDS_PH_PY118B.setValue("U_LineNum", i, i + 1)
'                    Next i
'
'                    oMat1.LoadFromDataSource
'                End If
'                Call PH_PY118_AddMatrixRow
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

Sub PH_PY118_AddMatrixRow()
    Dim oRow    As Long
    
    On Error GoTo PH_PY118_AddMatrixRow_Error
    
    Call oForm.Freeze(True)
    
'    '//[Mat1 용]
'    oMat1.FlushToDataSource
'    oRow = oMat1.VisualRowCount
'
'    If oMat1.VisualRowCount > 0 Then
'        If Trim(oDS_PH_PY118B.GetValue("U_FILD01", oRow - 1)) <> "" Then
'            If oDS_PH_PY118B.Size <= oMat1.VisualRowCount Then
'                oDS_PH_PY118B.InsertRecord (oRow)
'            End If
'            oDS_PH_PY118B.Offset = oRow
'            oDS_PH_PY118B.setValue "U_LineNum", oRow, oRow + 1
'            oDS_PH_PY118B.setValue "U_FILD01", oRow, ""
'            oDS_PH_PY118B.setValue "U_FILD02", oRow, ""
'            oDS_PH_PY118B.setValue "U_FILD03", oRow, 0
'            oMat1.LoadFromDataSource
'        Else
'            oDS_PH_PY118B.Offset = oRow - 1
'            oDS_PH_PY118B.setValue "U_LineNum", oRow - 1, oRow
'            oDS_PH_PY118B.setValue "U_FILD01", oRow - 1, ""
'            oDS_PH_PY118B.setValue "U_FILD02", oRow - 1, ""
'            oDS_PH_PY118B.setValue "U_FILD03", oRow - 1, 0
'            oMat1.LoadFromDataSource
'        End If
'    ElseIf oMat1.VisualRowCount = 0 Then
'        oDS_PH_PY118B.Offset = oRow
'        oDS_PH_PY118B.setValue "U_LineNum", oRow, oRow + 1
'        oDS_PH_PY118B.setValue "U_FILD01", oRow, ""
'        oDS_PH_PY118B.setValue "U_FILD02", oRow, ""
'        oDS_PH_PY118B.setValue "U_FILD03", oRow, 0
'        oMat1.LoadFromDataSource
'    End If
    
    Call oForm.Freeze(False)
    Exit Sub
PH_PY118_AddMatrixRow_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY118_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY118_FormClear()
On Error GoTo PH_PY118_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY118'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PH_PY118_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY118_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY118_DataValidCheck() As Boolean
On Error GoTo PH_PY118_DataValidCheck_Error
    PH_PY118_DataValidCheck = False
    Dim i       As Long
    Dim sQry    As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '//----------------------------------------------------------------------------------
    '//필수 체크
    '//----------------------------------------------------------------------------------
'    If Trim(oDS_PH_PY118A.GetValue("Code", 0)) = "" Then
'        Sbo_Application.SetStatusBarMessage "사원번호는 필수입니다.", bmt_Short, True
'        oForm.Items("Code").CLICK ct_Regular
'        Exit Function
'    End If
'
'    oMat1.FlushToDataSource
'    '// Matrix 마지막 행 삭제(DB 저장시)
'    If oDS_PH_PY118B.Size > 1 Then oDS_PH_PY118B.RemoveRecord (oDS_PH_PY118B.Size - 1)
'    oMat1.LoadFromDataSource
    
    PH_PY118_DataValidCheck = True
    
    Exit Function
    
    Set oRecordSet = Nothing
    
    
PH_PY118_DataValidCheck_Error:
    Set oRecordSet = Nothing
    PH_PY118_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY118_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Function PH_PY118_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY118_Validate_Error
    PH_PY118_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY118A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", bmt_Short, True
        PH_PY118_Validate = False
        GoTo PH_PY118_Validate_Exit
    End If
'
    If ValidateType = "수정" Then

    ElseIf ValidateType = "행삭제" Then

    ElseIf ValidateType = "취소" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY118_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY118_Validate_Error:
    PH_PY118_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY118_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function



Private Sub PH_PY118_Print_Report01()

    Dim DocNum          As String
    Dim ErrNum          As Integer
    Dim WinTitle        As String
    Dim ReportName      As String
    Dim sQry            As String
    
    Dim BPLID           As String
    Dim ItmBsort        As String
    Dim DocDate         As String
    
    Dim oRecordSet      As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY118_Print_Report01_Error
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '/ ODBC 연결 체크
    If ConnectODBC = False Then
        GoTo PH_PY118_Print_Report01_Error
    End If
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/

    WinTitle = "[S142] 발주서"
    ReportName = "S142_1.rpt"
    sQry = "EXEC PH_PY118_1 '" & oForm.Items("8").Specific.VALUE & "'"
    
    '/ Formula 수식필드
    ReDim gRpt_Formula(1)
    ReDim gRpt_Formula_Value(1)
    
    '/ SubReport
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)
    
    '/ Procedure 실행"
    sQry = "EXEC [PS_PP820_01] '" & BPLID & "', '" & ItmBsort & "', '" & DocDate & "'"
    
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
        If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "1", "Y", "V") = False Then
            Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
        End If
    Else
        Sbo_Application.SetStatusBarMessage "조회된 데이터가 없습니다.", bmt_Short, True
    End If
    
    Exit Sub
    
PH_PY118_Print_Report01_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY118_Print_Report01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub



Private Sub SetSendMail()
On Error GoTo Error_Message
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim ErrNum          As Integer
    
'/ Check
    ErrNum = 0
    sFrEmail = oForm.Items("FrEmail").Specific.String
    sFrSMTP = oForm.Items("FrSMTP").Specific.String
    sFrPWD = oForm.Items("FrPWD").Specific.String
    Select Case True
    Case Trim$(sFrEmail) = ""
        ErrNum = 1
        GoTo Error_Message
    Case Trim$(sFrSMTP) = ""
        ErrNum = 2
        GoTo Error_Message
    Case Trim$(sFrPWD) = ""
        ErrNum = 3
        GoTo Error_Message
    End Select
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    sQry = " SELECT U_FrSMTP, U_FrEMAIL, U_FrPWD FROM [@PH_PY118A] WHERE Code='1'"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
        sQry = "INSERT INTO [@PH_PY118A] (Code, Name, U_FrSMTP, U_FrEMAIL, U_FrPWD) values ('1','1','"
        sQry = sQry & Trim$(sFrSMTP) & "', '" & Trim$(sFrEmail) & "', '" & Trim$(sFrPWD) & "')"
        oRecordSet.DoQuery sQry
    Else
        sQry = "UPDATE  [@PH_PY118A] SET   U_FrSMTP = '" & Trim$(sFrSMTP) & "'"
        sQry = sQry & " , U_FrEMAIL = '" & Trim$(sFrEmail) & "'"
        sQry = sQry & " , U_FrPWD = '" & Trim$(sFrPWD) & "'"
        sQry = sQry & " WHERE Code  = '1'"
        oRecordSet.DoQuery sQry
    End If
    
    Set oRecordSet = Nothing
    Exit Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Error_Message:
    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "보내는 사람주소는 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 2 Then
        Sbo_Application.StatusBar.SetText "SMTP Server는 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 3 Then
        Sbo_Application.StatusBar.SetText "PassWord는 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
    Else
        Sbo_Application.StatusBar.SetText "SetSendMail Error :" & Space$(10) & Err.Description, bmt_Short, smt_Error
    End If
End Sub


Private Sub PrintChk()
    Dim i As Integer
    
    If oMat1.RowCount = 1 Then
        Exit Sub
    End If
    
    oMat1.FlushToDataSource
    For i = 0 To oDS_PH_PY118B.Size - 1
        oDS_PH_PY118B.Offset = i
        If oPrtChk = True Then
            oDS_PH_PY118B.setValue "U_Col11", i, "N"
        Else
            oDS_PH_PY118B.setValue "U_Col11", i, "Y"
        End If
    Next i
    oMat1.LoadFromDataSource
    If oPrtChk = True Then
        oPrtChk = False
    Else
        oPrtChk = True
    End If
End Sub
Private Sub eMail_Process()
On Error GoTo Error_Message
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim ErrNum          As Integer
    Dim cnt&, i&, oRow&
    Dim RetVal          As String
    Dim GNTSTR(1 To 9)  As String
'/ Check
    ErrNum = 0
    oMat1.FlushToDataSource
    sFrEmail = oForm.Items("FrEmail").Specific.String
    sFrSMTP = oForm.Items("FrSMTP").Specific.String
    sFrPWD = oForm.Items("FrPWD").Specific.String
    oREMARK = oForm.Items("Remark").Specific.String
    
    If InStr(1, sFrSMTP, ":", vbTextCompare) = 0 Then
        sFrSMTPSrv = sFrSMTP
        sFrSMTPPort = "25"
    Else
        sFrSMTPSrv = Left$(sFrSMTP, InStr(1, sFrSMTP, ":", vbTextCompare) - 1)
        sFrSMTPPort = Mid$(sFrSMTP, InStr(1, sFrSMTP, ":", vbTextCompare) + 1, Len(sFrSMTP) - InStr(1, sFrSMTP, ":", vbTextCompare))
    End If
    
    Select Case True
    Case oMat1.RowCount = 0
        ErrNum = 1
        GoTo Error_Message
    Case Trim$(sFrEmail) = ""
        ErrNum = 3
        GoTo Error_Message
    Case Trim$(sFrSMTP) = ""
        ErrNum = 4
        GoTo Error_Message
    Case Trim$(sFrPWD) = ""
        ErrNum = 5
        GoTo Error_Message
    End Select
'/ 초기화
    For i = 1 To 36
        ArrPayHead(i) = "---"
        ArrPayAmt(i) = 0
    Next i
    For i = 1 To 36
        ArrSubHead(i) = "---"
        ArrSubAmt(i) = 0
    Next i
    For i = 1 To 18
        ArrGntHead(i) = "---"
        ArrGntAmt(i) = 0
    Next i
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
 '/1. 수당/공제/근태 항목
 
    '//수당 항목
    sQry = "SELECT T0.U_CSUNAM"
    sQry = sQry & " FROM [@PH_PY102B] T0 INNER JOIN [@PH_PY102A] T1 ON T0.Code = T1.Code"
    sQry = sQry & " WHERE U_CLTCOD = '" & Trim$(oCLTCOD) & "'"
    sQry = sQry & " AND (T1.U_YM = '" & Trim$(oJOBYMM) & "' OR (T1.U_YM <> '" & Trim$(oJOBYMM) & "' AND T1.U_YM = (SELECT MAX(U_YM) FROM [@PH_PY102A] WHERE U_YM <= '" & Trim$(oJOBYMM) & "' )))"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
        ErrNum = 2
        GoTo Error_Message
    Else
        For i = 1 To 36
            If i <= oRecordSet.RecordCount Then
                ArrPayHead(i) = oRecordSet.Fields(0).VALUE
                oRecordSet.MoveNext
            End If
        Next i
    End If
    
    '//공제 항목
    sQry = "SELECT T0.U_CSUNAM"
    sQry = sQry & " FROM [@PH_PY103B] T0 INNER JOIN [@PH_PY103A] T1 ON T0.Code = T1.Code"
    sQry = sQry & " WHERE U_CLTCOD = '" & Trim$(oCLTCOD) & "'"
    sQry = sQry & " AND (T1.U_YM = '" & Trim$(oJOBYMM) & "' OR (T1.U_YM <> '" & Trim$(oJOBYMM) & "' AND T1.U_YM = (SELECT MAX(U_YM) FROM [@PH_PY103A] WHERE U_YM <= '" & Trim$(oJOBYMM) & "' )))"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
        ErrNum = 2
        GoTo Error_Message
    Else
        For i = 1 To 36
            If i <= oRecordSet.RecordCount Then
                ArrSubHead(i) = oRecordSet.Fields(0).VALUE
                oRecordSet.MoveNext
            End If
        Next i
    End If

'/ 2. 근태관련(1~14)
    ArrGntHead(1) = "근로일수":
    ArrGntHead(2) = "특근일수":
    ArrGntHead(3) = "유급일수"
    ArrGntHead(4) = "연장근로시간":
    ArrGntHead(5) = "야간근로":
    ArrGntHead(6) = "휴일기본"
    ArrGntHead(7) = "년월차":
    ArrGntHead(8) = "교대(주)":
    ArrGntHead(9) = "교대(야)"
    ArrGntHead(10) = "위해일수"
    ArrGntHead(11) = "생휴발생"
    ArrGntHead(12) = "생휴사용"
    ArrGntHead(13) = "지급율(%)"
    ArrGntHead(14) = "부양가족"
    ArrGntHead(15) = "일당액"
    ArrGntHead(16) = "---"
    ArrGntHead(17) = "---"
    ArrGntHead(18) = "---"
    
'    sQry = "SELECT CODE, U_FILNAM, U_FILD01 FROM [@ZPY231H] "
'    If Trim$(oJOBTYP) <> "2" Then
'        sQry = sQry & " WHERE CODE IN ('RG01','RG02','RG03','RG04','RG05','RG06','RG07','RG08','RG09') "
'    Else
'        sQry = sQry & " WHERE CODE IN ('RG21','RG22','RG23','RG24','RG25','RG26','RG27','RG28','RG29') "
'    End If
'    sQry = sQry & " ORDER BY CODE"
'    oRecordSet.DoQuery sQry
'    If oRecordSet.RecordCount = 0 Then
'        '// 상여일 경우 근태안쓰니 적용율 상여율 표시
'        If Trim$(oJOBTYP) = "2" Then
'            ArrGntHead(1) = "":  ArrGntHead(2) = "":        ArrGntHead(3) = ""
'            ArrGntHead(4) = "":  ArrGntHead(5) = "":        ArrGntHead(6) = ""
'            ArrGntHead(7) = "":  ArrGntHead(8) = "상여율":        ArrGntHead(9) = "적용율"
'        Else
'            ArrGntHead(1) = "근무시간":  ArrGntHead(2) = "연장시간":        ArrGntHead(3) = "야간시간"
'            ArrGntHead(4) = "특근시간":  ArrGntHead(5) = "특근연장":        ArrGntHead(6) = "특근야간"
'            ArrGntHead(7) = "근무일":  ArrGntHead(8) = "주휴일":        ArrGntHead(9) = "지.조.외"
'        End If
'
'        For i = 1 To 9
'            GNTSTR(i) = "0"
'        Next i
'    Else
'        Do Until oRecordSet.EOF
'            i = Mid$(oRecordSet.Fields(0).Value, 4, 1)
'            ArrGntHead(i) = oRecordSet.Fields(1).Value
'            GNTSTR(i) = oRecordSet.Fields(2).Value
'            oRecordSet.MoveNext
'        Loop
'    End If
'    ArrGntHead(10) = "---": ArrGntHead(11) = "---": ArrGntHead(12) = "---": ArrGntHead(13) = "---"
    
    
'/ 3. 전체 적용사항
    oCLTNAM = oCompany.CompanyName
    Select Case Trim$(oJOBTYP)
        Case "1": oPRTTIL = Left$(oJOBYMM, 4) & "년 " & Mid$(oJOBYMM, 5, 2) & "월 급여 명세서"
        Case "2": oPRTTIL = Left$(oJOBYMM, 4) & "년 " & Mid$(oJOBYMM, 5, 2) & "월 상여 명세서"
    End Select
    oPRTSUB = "(  " & oForm.Items("JOBGBN").Specific.Selected.Description & "  )"
'/ 4.사원별 DM발송
    cnt = 0
    For oRow = 0 To oDS_PH_PY118B.Size - 1
        oDS_PH_PY118B.Offset = oRow
        If Trim$(oDS_PH_PY118B.GetValue("U_Col08", oRow)) = "Y" Then
            sMSTCOD = oDS_PH_PY118B.GetValue("U_MSTCOD", oRow)
            sMSTNAM = oDS_PH_PY118B.GetValue("U_MSTNAM", oRow)
            sDPTNAM = oDS_PH_PY118B.GetValue("U_Col01", oRow)
            sPOSITION = oDS_PH_PY118B.GetValue("U_Col02", oRow)
            sTOEmail = oDS_PH_PY118B.GetValue("U_Col06", oRow)
            
            If Trim$(sTOEmail) = "" Then
                oDS_PH_PY118B.setValue "U_Col07", oRow, "메일주소누락"
            Else
                '/ 급여정보 가져오기
            
                sQry = "SELECT  T0.U_MSTCOD,"
                sQry = sQry & " T0.U_CSUD01, T0.U_CSUD02, T0.U_CSUD03, T0.U_CSUD04, T0.U_CSUD05, T0.U_CSUD06, T0.U_CSUD07, T0.U_CSUD08,T0.U_CSUD09,"
                sQry = sQry & " T0.U_CSUD10, T0.U_CSUD11, T0.U_CSUD12, T0.U_CSUD13, T0.U_CSUD14, T0.U_CSUD15, T0.U_CSUD16,T0.U_CSUD17,T0.U_CSUD18,"
                sQry = sQry & " T0.U_CSUD19, T0.U_CSUD20, T0.U_CSUD21, T0.U_CSUD22, T0.U_CSUD23, T0.U_CSUD24, T0.U_CSUD25, T0.U_CSUD26, T0.U_CSUD27,"
                sQry = sQry & " T0.U_CSUD28, T0.U_CSUD29, T0.U_CSUD30, T0.U_CSUD31, T0.U_CSUD32, T0.U_CSUD33, T0.U_CSUD34, T0.U_CSUD35, T0.U_CSUD36,"
                sQry = sQry & " T0.U_GONG01, T0.U_GONG02, T0.U_GONG03, T0.U_GONG04, T0.U_GONG05, T0.U_GONG06, T0.U_GONG07, T0.U_GONG08, T0.U_GONG09,"
                sQry = sQry & " T0.U_GONG10, T0.U_GONG11, T0.U_GONG12, T0.U_GONG13, T0.U_GONG14, T0.U_GONG15, T0.U_GONG16, T0.U_GONG17, T0.U_GONG18,"
                sQry = sQry & " T0.U_GONG19, T0.U_GONG20, T0.U_GONG21, T0.U_GONG22, T0.U_GONG23, T0.U_GONG24, T0.U_GONG25, T0.U_GONG26, T0.U_GONG27,"
                sQry = sQry & " T0.U_GONG28, T0.U_GONG29, T0.U_GONG30, T0.U_GONG31, T0.U_GONG32, T0.U_GONG33, T0.U_GONG34, T0.U_GONG35, T0.U_GONG36,"
                sQry = sQry & " T0.U_TOTPAY, T0.U_TOTGON, T0.U_SILJIG, T0.U_CLTNAM,"
                sQry = sQry & " T3.U_GetDay, T3.U_WoHDay, T3.U_PayDay, T3.U_Extend, T3.U_Midnight, T3.U_Special, T3.U_YCHHGA,"
                sQry = sQry & " T3.U_EtcDAY1 , T3.U_EtcDAY2, T3.U_WHMDAY, T3.U_SNHDAY, T3.U_SNHHGA, T0.U_APPRAT, T0.U_BUYNSU"
                sQry = sQry & " FROM [@PH_PY112A] T0"
                sQry = sQry & " LEFT JOIN ( SELECT  T2.U_MSTCOD, T1.U_YM, T2.U_GetDay, T2.U_WoHDay, T2.U_PayDay, T2.U_Extend, T2.U_Midnight, T2.U_Special,"
                sQry = sQry & "                     T2.U_YCHHGA, T2.U_EtcDAY1 , T2.U_EtcDAY2, T2.U_WHMDAY, T2.U_SNHDAY, T2.U_SNHHGA"
                sQry = sQry & "             FROM [@PH_PY017B] T2 INNER JOIN [@PH_PY017A] T1 ON T2.Code = T1.Code"
                sQry = sQry & "           ) T3 ON T0.U_YM = T3.U_YM AND T0.U_MSTCOD = T3.U_MSTCOD"
                sQry = sQry & " WHERE   T0.U_YM = '" & Trim$(oJOBYMM) & "'"
                sQry = sQry & " AND     T0.U_JOBTYP = '" & Trim$(oJOBTYP) & "'"
                sQry = sQry & " AND     T0.U_JOBGBN = '" & Trim$(oJOBGBN) & "'"
                sQry = sQry & " AND     (T0.U_JOBTRG = '" & Trim$(oPAYSEL) & "' OR ( T0.U_JOBTRG <> '" & Trim$(oPAYSEL) & "' AND T0.U_JOBTRG LIKE '" & Trim$(oPAYSEL) & "'))"
                sQry = sQry & " AND     T0.U_MSTCOD = '" & Trim$(sMSTCOD) & "'"
                
                oRecordSet.DoQuery sQry
                If oRecordSet.RecordCount > 0 Then
                    For i = 1 To 36
                        If oRecordSet.Fields("U_CSUD" & Format(i, "00")).VALUE = Null Then oRecordSet.Fields("U_CSUD" & Format(i, "00")).VALUE = 0
                        ArrPayAmt(i) = oRecordSet.Fields("U_CSUD" & Format(i, "00")).VALUE
                    Next i
                    For i = 1 To 36
                        If oRecordSet.Fields("U_GONG" & Format(i, "00")).VALUE = Null Then oRecordSet.Fields("U_GONG" & Format(i, "00")).VALUE = 0
                        ArrSubAmt(i) = oRecordSet.Fields("U_GONG" & Format(i, "00")).VALUE
                    Next i
                    For i = 1 To 14
                        If oRecordSet.Fields(i + 76).VALUE = Null Then oRecordSet.Fields(i + 76).VALUE = 0
                        ArrGntAmt(i) = oRecordSet.Fields(i + 76).VALUE
                    Next i
                    
                    sTOTPAY = oRecordSet.Fields("U_TOTPAY").VALUE
                    sTOTGON = oRecordSet.Fields("U_TOTGON").VALUE
                    sSILJIG = oRecordSet.Fields("U_SILJIG").VALUE
                    If oRecordSet.Fields("U_CLTNAM").VALUE <> "" Then
                        oCLTNAM = oRecordSet.Fields("U_CLTNAM").VALUE
                    End If
                
            
                Else
                    oDS_PH_PY118B.setValue "U_Col07", oRow, "자료누락."
                    For i = 1 To 36
                        ArrPayAmt(i) = 0
                    Next i
                    For i = 1 To 36
                        ArrSubAmt(i) = 0
                    Next i
                    For i = 1 To 13
                        ArrGntAmt(i) = 0
                    Next i
                End If
            End If
            
            Call Create_Html2
            RetVal = Send_eMail
            If Trim$(RetVal) = "True" Then
                cnt = cnt + 1
                oDS_PH_PY118B.setValue "U_Col07", oRow, "Success"
            Else
                oDS_PH_PY118B.setValue "U_Col07", oRow, Left$("Failure:" & RetVal, 50)
            End If
            oMat1.SetLineData oRow + 1
        End If
    Next oRow
 '/ End
    Set oRecordSet = Nothing
    Exit Sub
'/////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "선택된 자료가 없습니다.", bmt_Short, smt_Error
    ElseIf ErrNum = 2 Then
        Sbo_Application.StatusBar.SetText "조건과 일치하는 자료가 없습니다.", bmt_Short, smt_Error
    ElseIf ErrNum = 3 Then
        Sbo_Application.StatusBar.SetText "보내는 사람 주소는 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 4 Then
        Sbo_Application.StatusBar.SetText "SMTP Server는 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 5 Then
        Sbo_Application.StatusBar.SetText "Password는 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
    Else
        Sbo_Application.StatusBar.SetText "Send_eMail Error: " & Space$(10) & Err.Description, bmt_Short, smt_Error
    End If
End Sub

Private Function Send_eMail() As String
On Error GoTo Error_Message

    Dim ErrNum          As Integer
    Dim MailSender  As CDO.message          '// CDO.Message 개체에 대한 참조를 작성하고 반환
    Dim iConf       As CDO.Configuration    '// CDO.Configuration개체에 대한 참조를 작성하고 반환
    ErrNum = 0
    
    Set MailSender = CreateObject("CDO.Message")
    Set iConf = CreateObject("CDO.Configuration")
    
    '2011.12.22 경보제약 메일발송 익명인증사용때문에 조건 추가
    If oForm.Items("AUTCHK").Specific.Checked = False Then
        
        With iConf.Fields
        
            .Refresh
            .Item(cdoSendUsingMethod) = cdoSendUsingPort   '/1:로컬smtp로 메일전송, 2 cdoSendUsingPort: 외부smtp로 메일전송
            .Item(cdoSMTPServer) = Trim$(sFrSMTPSrv) '"mail.emdc.co.kr","mail.care-line.co.kr" '"보내는 사람의 SMTP Server Name" '(ex: mail.xxx.com)
            .Item(cdoSMTPConnectionTimeout) = 10
            .Item(cdoSMTPAuthenticate) = cdoBasic '기본인증
            .Item(cdoSendUserName) = Trim$(sFrEmail)  '"hammi97@emdc.co.kr""sap2@care-line.co.kr" '"보내는사람주소" '(ex : peter@xxx.com)
            .Item(cdoSendPassword) = Trim$(sFrPWD) 'mi0215" ''"sap02" '"보내는 사람의 Password"
            .Item(cdoURLGetLatestVersion) = True
            .Item(cdoSMTPServerPort) = sFrSMTPPort
            '.Item(cdoSMTPServerPort) = 25 '/통상 25번포트 네이버 pop3:110포트
    
            .Update
                    
        End With
        
    Else
        
        With iConf.Fields
        
            .Refresh
            .Item(cdoSendUsingMethod) = cdoSendUsingPort   '/1:로컬smtp로 메일전송, 2 cdoSendUsingPort: 외부smtp로 메일전송
            .Item(cdoSMTPServer) = Trim$(sFrSMTPSrv) '"mail.care-line.co.kr" '"보내는 사람의 SMTP Server Name" '(ex: mail.xxx.com)
            .Item(cdoSMTPConnectionTimeout) = 10
            .Item(cdoSMTPAuthenticate) = cdoAnonymous '익명인증
            .Item(cdoURLGetLatestVersion) = True
            .Item(cdoSMTPServerPort) = sFrSMTPPort
            '.Item(cdoSMTPServerPort) = 25 '/통상 25번포트 네이버 pop3:110포트
    
            .Update
                    
        End With
        
    End If
                
    With MailSender
         .Configuration = iConf
         .From = Trim$(sFrEmail) '전송자/수신자 이름만 넘김
         .To = Trim$(sTOEmail)
         .Subject = oPRTTIL
         .HTMLBody = sHtml
        ' .TextBody = "귀하의 노고에 감사드립니다."
         .BodyPart.Charset = "ks_c_5601-1987"
         .HTMLBodyPart.Charset = "ks_c_5601-1987"
         .send
        '.AddAttachment "C:\files\mybook.doc"   '/ 첨부파일
    End With

    Set MailSender = Nothing
    Set iConf = Nothing
 '/ End
    Sbo_Application.StatusBar.SetText "작업을 완료하였습니다.", bmt_Short, smt_Success
    Send_eMail = "True"
    '-2147220973
    Exit Function
'/////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set MailSender = Nothing
    Set iConf = Nothing
    Sbo_Application.StatusBar.SetText Err.Number & " : " & Err.Description, bmt_Short, smt_Error
    Send_eMail = Err.Number & " : " & Err.Description
End Function

Private Sub Create_Html2()
   Dim i As Integer
   
   sHtml = "<html>" & vbCrLf
   sHtml = sHtml & "<head>" & vbCrLf
   sHtml = sHtml & "<title>급여명세서 e-MAIL발송</title>" & vbCrLf
   sHtml = sHtml & "<style type=""text/css"">" & vbCrLf
   sHtml = sHtml & "<!--" & vbCrLf
   sHtml = sHtml & " td {  font-size: 9pt; line-height: 14pt; color: #000000}" & vbCrLf
   sHtml = sHtml & ".목록 {  font-size: 9pt; font-weight: bold; color: #FFFFFF}" & vbCrLf
   sHtml = sHtml & ".마침글 {  font-size: 9pt; color: #000000; font-weight: bold}}" & vbCrLf '파란색:#3333FF
   sHtml = sHtml & "-->" & vbCrLf
   sHtml = sHtml & "</style>" & vbCrLf
   sHtml = sHtml & "</head>" & vbCrLf
   sHtml = sHtml & "<body>" & vbCrLf
   sHtml = sHtml & "</style>" & vbCrLf
   sHtml = sHtml & "</head>" & vbCrLf
   sHtml = sHtml & "<body>" & vbCrLf
'/ BODY /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
   '// 수정0->1
   sHtml = sHtml & " <table border=""0"" cellspacing=""0"" cellpadding=""0""  align=""center"" width=""900"">" & vbCrLf
   sHtml = sHtml & "      <tr></tr>" & vbCrLf
'///**********************************************************************************************************************/
'/ 타이틀정보
'///**********************************************************************************************************************/
   sHtml = sHtml & "     <table border=""0"" bgcolor=""white"" cellspacing=""0"" cellpadding=""2"" width=""800"">" & vbCrLf
   sHtml = sHtml & "                      <tr bgcolor=""white"" align=""Center"">" & vbCrLf
   sHtml = sHtml & "                          <td width=""100%""><h4><b>&lt;" & Trim$(oPRTTIL) & "&gt;</h></td>" & vbCrLf
   sHtml = sHtml & "                      </tr>" & vbCrLf
   sHtml = sHtml & "                      <tr bgcolor=""white"" align=""Center"">" & vbCrLf
   sHtml = sHtml & "                          <td width=""100%""><h4><b>" & Trim$(oPRTSUB) & "</h></td>" & vbCrLf
   sHtml = sHtml & "                      </tr>" & vbCrLf
   sHtml = sHtml & "     </table><br>" & vbCrLf
'///**********************************************************************************************************************/
'/ 사원정보
'///**********************************************************************************************************************/
   sHtml = sHtml & "     <table border=""1"" bgcolor=""white"" cellspacing=""0"" cellpadding=""2"" width=""800"">" & vbCrLf
   sHtml = sHtml & "                      <tr bgcolor=""RGB(239,235,222)"" align=""center"">" & vbCrLf
   sHtml = sHtml & "                          <td width=""12%"">사  번</td>" & vbCrLf
   sHtml = sHtml & "                          <td width=""13%"">성  명</td>" & vbCrLf
   sHtml = sHtml & "                          <td width=""15%"">부  서</td>" & vbCrLf
   sHtml = sHtml & "                          <td width=""15%"">직  책</td>" & vbCrLf
   sHtml = sHtml & "                          <td width=""15%"">지급총액</td>" & vbCrLf
   sHtml = sHtml & "                          <td width=""15%"">공제총액</td>" & vbCrLf
   sHtml = sHtml & "                          <td width=""15%"">실지급액</td>" & vbCrLf
   sHtml = sHtml & "                      </tr>" & vbCrLf
   sHtml = sHtml & "                    <tr bgcolor=""white"" align=""CENTER"">" & vbCrLf
   sHtml = sHtml & "                        <td width=""12%"" align=""CENTER"">" & sMSTCOD & "</td>" & vbCrLf
   sHtml = sHtml & "                        <td width=""13%"" align=""CENTER"">" & sMSTNAM & "</td>" & vbCrLf
   sHtml = sHtml & "                        <td width=""15%"" align=""CENTER"">" & sDPTNAM & "</td>" & vbCrLf
   sHtml = sHtml & "                        <td width=""15%"" align=""CENTER"">" & sPOSITION & "</td>" & vbCrLf
   sHtml = sHtml & "                        <td width=""15%"" align=""Right""><b>" & Format$(sTOTPAY, "#,###,##0") & "</b></td>" & vbCrLf
   sHtml = sHtml & "                        <td width=""15%"" align=""Right""><b>" & Format$(sTOTGON, "#,###,##0") & "</b></td>" & vbCrLf
   sHtml = sHtml & "                        <td width=""15%"" align=""Right""><b>" & Format$(sSILJIG, "#,###,##0") & "</b></td>" & vbCrLf
   sHtml = sHtml & "                    </tr>" & vbCrLf
   sHtml = sHtml & "     </table><br>" & vbCrLf
   sHtml = sHtml & "  <tr>" & vbCrLf
'///**********************************************************************************************************************/
'/ 내역
'///**********************************************************************************************************************/
   sHtml = sHtml & "                <table border=""1""  bgcolor=""white"" cellspacing=""0"" cellpadding=""2"" width=""800"">" & vbCrLf
   
   sHtml = sHtml & "                <!--- 근태항목 ----------------------------------------------------------->" & vbCrLf
   sHtml = sHtml & "                    <tr bgcolor=""RGB(239,235,222)"" align=""center"">" & vbCrLf
   sHtml = sHtml & "                        <td width=""10%"" rowspan=""4"">근태항목</td>" & vbCrLf
   '/ 근태항목(01~09항목)
   For i = 1 To 9
        sHtml = sHtml & "                        <td  width=""10%"">" & IIf(Trim$(ArrGntHead(i)) = Space$(0), "-", ArrGntHead(i)) & "</td>" & vbCrLf
   Next i
   sHtml = sHtml & "                    </tr>" & vbCrLf
   sHtml = sHtml & "                    <tr bgcolor=""white"" align=""Right"">" & vbCrLf
   For i = 1 To 9
        sHtml = sHtml & "                        <td width=""10%"">" & ArrGntAmt(i) & "</td>" & vbCrLf
   Next i
   sHtml = sHtml & "                    </tr>" & vbCrLf
   
   '/ 근태항목(10~18항목)
   sHtml = sHtml & "                    <tr bgcolor=""RGB(239,235,222)"" align=""center"">" & vbCrLf
   For i = 10 To 18
        sHtml = sHtml & "                        <td  width=""10%"">" & IIf(Trim$(ArrGntHead(i)) = Space$(0), "-", ArrGntHead(i)) & "</td>" & vbCrLf
   Next i
   sHtml = sHtml & "                    </tr>" & vbCrLf
   sHtml = sHtml & "                    <tr bgcolor=""white"" align=""Right"">" & vbCrLf
   For i = 10 To 18
        sHtml = sHtml & "                        <td width=""10%"">" & ArrGntAmt(i) & "</td>" & vbCrLf
   Next i
   sHtml = sHtml & "                    </tr>" & vbCrLf
   
   
   sHtml = sHtml & "                <!--- 지급항목 ----------------------------------------------------------->" & vbCrLf
   sHtml = sHtml & "                    <tr bgcolor=""RGB(239,235,222)"" align=""center"">" & vbCrLf
   sHtml = sHtml & "                        <td  width=""10%"" rowspan=""8"">지급항목</td>" & vbCrLf
'/ 지급항목(01~09항목)
   For i = 1 To 9
        sHtml = sHtml & "                        <td  width=""10%"">" & ArrPayHead(i) & "</td>" & vbCrLf
   Next i
   sHtml = sHtml & "                    </tr>" & vbCrLf
   sHtml = sHtml & "                    <tr bgcolor=""white"" align=""Right"">" & vbCrLf
   For i = 1 To 9
        sHtml = sHtml & "                        <td width=""10%"">" & Format$(ArrPayAmt(i), "#,###,##0") & "</td>" & vbCrLf
   Next i
   sHtml = sHtml & "                    </tr>" & vbCrLf
'/ 지급항목(10~18항목)
   sHtml = sHtml & "                    <tr bgcolor=""RGB(239,235,222)"" align=""center"">" & vbCrLf
   For i = 10 To 18
        sHtml = sHtml & "                        <td width=""10%"">" & IIf(Trim$(ArrPayHead(i)) = Space$(0), "-", ArrPayHead(i)) & "</td>" & vbCrLf
   Next i
   sHtml = sHtml & "                    </tr>" & vbCrLf
   sHtml = sHtml & "                    <tr bgcolor=""white"" align=""Right"">" & vbCrLf
   For i = 10 To 18
   sHtml = sHtml & "                        <td width=""10%"">" & Format$(ArrPayAmt(i), "#,###,##0") & "</td>" & vbCrLf
   Next i
   sHtml = sHtml & "                    </tr>" & vbCrLf
'/ 지급항목(19~27항목)
   sHtml = sHtml & "                    <tr bgcolor=""RGB(239,235,222)"" align=""center"">" & vbCrLf
   For i = 19 To 27
   sHtml = sHtml & "                        <td width=""10%"">" & IIf(Trim$(ArrPayHead(i)) = Space$(0), "-", ArrPayHead(i)) & "</td>" & vbCrLf
   Next i
   sHtml = sHtml & "                    </tr>" & vbCrLf
   sHtml = sHtml & "                    <tr bgcolor=""white"" align=""Right"">" & vbCrLf
   For i = 19 To 27
   sHtml = sHtml & "                        <td width=""10%"">" & Format$(ArrPayAmt(i), "#,###,##0") & "</td>" & vbCrLf
   Next i
   sHtml = sHtml & "                    </tr>" & vbCrLf
'/ 지급항목(28~36항목)
   sHtml = sHtml & "                    <tr bgcolor=""RGB(239,235,222)"" align=""center"">" & vbCrLf
   For i = 28 To 36
   sHtml = sHtml & "                        <td width=""10%"">" & IIf(Trim$(ArrPayHead(i)) = Space$(0), "-", ArrPayHead(i)) & "</td>" & vbCrLf
   Next i
   sHtml = sHtml & "                    </tr>" & vbCrLf
   sHtml = sHtml & "                    <tr bgcolor=""white"" align=""Right"">" & vbCrLf
   For i = 28 To 36
   sHtml = sHtml & "                        <td width=""10%"">" & Format$(ArrPayAmt(i), "#,###,##0") & "</td>" & vbCrLf
   Next i
   sHtml = sHtml & "                    </tr>" & vbCrLf

   sHtml = sHtml & "<!--- 공제항목 ----------------------------------------------------------->" & vbCrLf
   sHtml = sHtml & "                    <tr bgcolor=""RGB(239,235,222)"" align=""center"">" & vbCrLf
   sHtml = sHtml & "                        <td width=""10%"" rowspan=""8"">공제항목</td>" & vbCrLf
 '/ 공제항목(01-09항목)
   For i = 1 To 9
        sHtml = sHtml & "                        <td  width=""10%"">" & IIf(Trim$(ArrSubHead(i)) = Space$(0), "-", ArrSubHead(i)) & "</td>" & vbCrLf
   Next i
   sHtml = sHtml & "                    </tr>" & vbCrLf
   sHtml = sHtml & "                    <tr bgcolor=""white"" align=""Right"">" & vbCrLf
   For i = 1 To 9
        sHtml = sHtml & "                        <td width=""10%"">" & Format$(ArrSubAmt(i), "#,###,##0") & "</td>" & vbCrLf
   Next i
   sHtml = sHtml & "                    </tr>" & vbCrLf
 '/ 공제항목(08-14항목)
   sHtml = sHtml & "                    <tr bgcolor=""RGB(239,235,222)"" align=""center"">" & vbCrLf
   For i = 10 To 18  '14
        sHtml = sHtml & "                        <td width=""10%"">" & IIf(Trim$(ArrSubHead(i)) = Space$(0), "-", ArrSubHead(i)) & "</td>" & vbCrLf
   Next i
   sHtml = sHtml & "                    </tr>" & vbCrLf
   sHtml = sHtml & "                    <tr bgcolor=""white"" align=""Right"">" & vbCrLf
   For i = 10 To 18  '14
        sHtml = sHtml & "                        <td width=""10%"">" & Format$(ArrSubAmt(i), "#,###,##0") & "</td>" & vbCrLf
   Next i
   sHtml = sHtml & "                    </tr>" & vbCrLf
 '/ 공제항목(08-14항목)
   sHtml = sHtml & "                    <tr bgcolor=""RGB(239,235,222)"" align=""center"">" & vbCrLf
   For i = 19 To 27
        sHtml = sHtml & "                        <td width=""10%"">" & IIf(Trim$(ArrSubHead(i)) = Space$(0), "-", ArrSubHead(i)) & "</td>" & vbCrLf
   Next i
   sHtml = sHtml & "                    </tr>" & vbCrLf
   sHtml = sHtml & "                    <tr bgcolor=""white"" align=""Right"">" & vbCrLf
   For i = 19 To 27
        sHtml = sHtml & "                        <td width=""10%"">" & Format$(ArrSubAmt(i), "#,###,##0") & "</td>" & vbCrLf
   Next i
   sHtml = sHtml & "                    </tr>" & vbCrLf
   '/ 공제항목(08-14항목)
   sHtml = sHtml & "                    <tr bgcolor=""RGB(239,235,222)"" align=""center"">" & vbCrLf
   For i = 28 To 36  '14
        sHtml = sHtml & "                        <td width=""10%"">" & IIf(Trim$(ArrSubHead(i)) = Space$(0), "-", ArrSubHead(i)) & "</td>" & vbCrLf
   Next i
   sHtml = sHtml & "                    </tr>" & vbCrLf
   sHtml = sHtml & "                    <tr bgcolor=""white"" align=""Right"">" & vbCrLf
   For i = 28 To 36  '14
        sHtml = sHtml & "                        <td width=""10%"">" & Format$(ArrSubAmt(i), "#,###,##0") & "</td>" & vbCrLf
   Next i
   sHtml = sHtml & "                    </tr>" & vbCrLf
   

   sHtml = sHtml & "                </table>" & vbCrLf
'///**********************************************************************************************************************/
'/ 비고내용
'///**********************************************************************************************************************/
   sHtml = sHtml & "     <table border=""0"" bgcolor=""white"" cellspacing=""0"" cellpadding=""2"" width=""800"">" & vbCrLf
   sHtml = sHtml & "                      <tr bgcolor=""white"" align=""Left"">" & vbCrLf
   sHtml = sHtml & "                          <td width=""100%"">" & Trim$(oREMARK) & "</td>" & vbCrLf
   sHtml = sHtml & "                      </tr>" & vbCrLf
   sHtml = sHtml & "     </table><br>" & vbCrLf
   
'/
   sHtml = sHtml & "<!--- 지급항목테이블 끝. ----------------------------------------------------------------------->" & vbCrLf
   sHtml = sHtml & "      </td></tr>" & vbCrLf
   sHtml = sHtml & "      <tr>" & vbCrLf
   sHtml = sHtml & "       <td align=""left"" ><b class=""마침글"">" & Trim$(oCLTNAM) & "</b></td>" & vbCrLf
   sHtml = sHtml & "      </tr>" & vbCrLf
   sHtml = sHtml & " </table>" & vbCrLf
   
'/ END/~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
   sHtml = sHtml & "<!--- 본문 끝입니다. -------------------------------------------------------------------------->" & vbCrLf
   sHtml = sHtml & "</body>" & vbCrLf
   sHtml = sHtml & "</html>" & vbCrLf
'/
End Sub

Private Function Execution_Process() As Boolean
On Error GoTo Error_Message
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim ErrNum          As Integer
    Dim i&
    Dim V_StatusCnt&, oProValue&, TOTCNT&, tRow&  '//progbar
    
    
    ErrNum = 0
    '/ Check
    oJOBYMM = oForm.Items("JOBYMM").Specific.VALUE
    oJOBTYP = oForm.Items("JOBTYP").Specific.Selected.VALUE
    oJOBGBN = oForm.Items("JOBGBN").Specific.Selected.VALUE
    oPAYSEL = oForm.Items("PAYSEL").Specific.Selected.VALUE
    oCLTCOD = oForm.Items("CLTCOD").Specific.Selected.VALUE
    oMSTDPT = oForm.Items("MSTDPT").Specific.Selected.VALUE
    oMSTCOD = oForm.Items("MSTCOD").Specific.String

    If Trim$(oMSTCOD) = "" Then oMSTCOD = "%"
    Select Case True
    Case Trim$(oJOBYMM) = ""
        ErrNum = 1
        GoTo Error_Message
    End Select
    
   
    
'/ 초기화
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    oDS_PH_PY118B.Clear
    oMat1.LoadFromDataSource
    i = 0
    sQry = "SELECT T0.U_MSTCOD,T0.U_MSTNAM, T0.U_EmpID,  T2.U_CodeNM, T3.Name,"
    sQry = sQry & " ISNULL(CONVERT(CHAR(10), T1.U_StartDat, 20),'') AS U_INPDAT,"
    sQry = sQry & " ISNULL(CONVERT(CHAR(10),T1.U_TermDate, 20), '') AS U_OUTDAT, T0.U_SILJIG, T1.U_email"
    sQry = sQry & " FROM [@PH_PY112A] T0  INNER JOIN [@PH_PY001A] T1 ON T0.U_MSTCOD = T1.Code"
    sQry = sQry & " INNER JOIN [@PS_HR200L] T2 ON T1.U_TeamCode = T2.U_Code AND T2.Code = '1'"
    sQry = sQry & " INNER JOIN [OHPS] T3 ON T1.U_Position = T3.posID"
    sQry = sQry & " WHERE   T0.U_YM = '" & Trim$(oJOBYMM) & "'"
    sQry = sQry & " AND     T0.U_JOBTYP = '" & Trim$(oJOBTYP) & "'"
    sQry = sQry & " AND     T0.U_JOBGBN = '" & Trim$(oJOBGBN) & "'"
    sQry = sQry & " AND     (T1.U_PAYSEL = '" & Trim$(oPAYSEL) & "' OR T1.U_PAYSEL LIKE '" & Trim$(oPAYSEL) & "')"
    sQry = sQry & " AND     T0.U_CLTCOD = '" & Trim$(oCLTCOD) & "'"
    sQry = sQry & " AND     (T1.U_TeamCode = '" & Trim$(oMSTDPT) & "' OR T1.U_TeamCode LIKE '" & Trim$(oMSTDPT) & "')"
    sQry = sQry & " AND     (T1.Code = '" & Trim$(oMSTCOD) & "' OR T1.Code LIKE '" & Trim$(oMSTCOD) & "')"
    sQry = sQry & " ORDER BY T0.U_CLTCOD,  T0.U_TeamCode, T1.U_Position, T0.U_MSTCOD"
    
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
        ErrNum = 2
        GoTo Error_Message
    End If
    
    If Not oProgBar Is Nothing Then
        oProgBar.Stop
        Set oProgBar = Nothing
    End If
    '프로그레스 바    ///////////////////////////////////////
    Set oProgBar = Sbo_Application.StatusBar.CreateProgressBar("데이터 읽는중...!", 50, False)
    
    '최대값 구하기
    TOTCNT = oRecordSet.RecordCount
   
    V_StatusCnt = Round(TOTCNT / 50, 0)
    oProValue = 1
    tRow = 1
    '/////////////////////////////////////////////////////
    
    Do Until oRecordSet.EOF
        oDS_PH_PY118B.InsertRecord (i)
        oDS_PH_PY118B.Offset = i
        oDS_PH_PY118B.setValue "U_LineNum", i, i + 1
        oDS_PH_PY118B.setValue "U_MSTCOD", i, oRecordSet.Fields(0).VALUE
        oDS_PH_PY118B.setValue "U_MSTNAM", i, oRecordSet.Fields(1).VALUE
        oDS_PH_PY118B.setValue "U_EMPID", i, oRecordSet.Fields(2).VALUE
        oDS_PH_PY118B.setValue "U_Col01", i, oRecordSet.Fields(3).VALUE
        oDS_PH_PY118B.setValue "U_Col02", i, oRecordSet.Fields(4).VALUE
        oDS_PH_PY118B.setValue "U_Col03", i, oRecordSet.Fields(5).VALUE
        oDS_PH_PY118B.setValue "U_Col04", i, oRecordSet.Fields(6).VALUE
        oDS_PH_PY118B.setValue "U_Col05", i, Format$(oRecordSet.Fields(7).VALUE, "#,###,###,##0")
        oDS_PH_PY118B.setValue "U_Col06", i, oRecordSet.Fields(8).VALUE
        If Trim$(oRecordSet.Fields(8).VALUE) <> "" Then
            oDS_PH_PY118B.setValue "U_Col08", i, "Y"
        Else
            oDS_PH_PY118B.setValue "U_Col08", i, "N"
        End If
        i = i + 1
        oRecordSet.MoveNext
        
        If (TOTCNT > 50 And tRow = oProValue * V_StatusCnt) Or TOTCNT <= 50 Then
            oProgBar.Text = tRow & "/ " & TOTCNT & " 건 처리중...!"
            oProValue = oProValue + 1
            oProgBar.VALUE = oProValue
        End If
        tRow = tRow + 1
    Loop
    oPrtChk = True
    oProgBar.Stop
    oMat1.LoadFromDataSource
    
 '/ End
    Sbo_Application.StatusBar.SetText "작업을 완료하였습니다.", bmt_Short, smt_Success
    Execution_Process = True
    Set oRecordSet = Nothing
    Set oProgBar = Nothing
    Exit Function
'/////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set oRecordSet = Nothing
    If Not oProgBar Is Nothing Then
        oProgBar.Stop
        Set oProgBar = Nothing
    End If
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "귀속연월을 입력하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 2 Then
        Sbo_Application.StatusBar.SetText "조건과 일치하는 자료가 없습니다.", bmt_Short, smt_Error
    Else
        Sbo_Application.StatusBar.SetText "Execution_Process 실행 중 오류가 발생했습니다." & Space$(10) & Err.Description, bmt_Short, smt_Error
    End If
    Execution_Process = False
End Function




