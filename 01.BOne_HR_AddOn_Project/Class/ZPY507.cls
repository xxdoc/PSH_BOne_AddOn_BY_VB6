VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ZPY507"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//  SAP MANAGE UI API 2004 SDK Sample
'//****************************************************************************
'//  File           : ZPY507.cls
'//  Module         : 인사관리>원천징수>근로소득
'//  Desc           : 정산 결과 조회(전체)
'//  FormType       : 2010110507
'//  Create Date    : 2009.12.13
'//  Modified Date  :
'//  Creator        : Choi Dong Kwon
'//  Modifier       :
'//  Copyright  (c) Morning Data
'//****************************************************************************
Option Explicit
Public oFormUniqueID    As String
Public oForm            As SAPbouiCOM.Form
Private oGrid1          As SAPbouiCOM.Grid
Private oDS_ZPY507      As SAPbouiCOM.DataTable

Private Sub titleSetting()
On Error GoTo Error_Message
    Dim sQry            As String
    Dim iCol            As Long
    
    Dim COLNAM(150)     As String
    Dim oEditCol        As SAPbouiCOM.EditTextColumn
    Dim oComboCol       As SAPbouiCOM.ComboBoxColumn
    Dim oColumn         As SAPbouiCOM.GridColumn
    Dim oRecordSet      As SAPbobsCOM.Recordset
    
'/ Initial
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    oForm.Freeze True
    
    COLNAM(0) = "정산년도"
    COLNAM(1) = "정산구분"
    COLNAM(2) = "신고연월"
    COLNAM(3) = "지급일자"
    COLNAM(4) = "사업장"
    COLNAM(5) = "사원번호"
    COLNAM(6) = "사원순번"
    COLNAM(7) = "사원명"
    COLNAM(8) = "월별자료No"
    COLNAM(9) = "소득항목No"
    COLNAM(10) = "종전문서No"
    COLNAM(11) = "세액계산No"
    COLNAM(12) = "의료비No"
    COLNAM(13) = "기부금No"
    COLNAM(14) = "연금.저축No"
    COLNAM(15) = "현근무지총계"
    COLNAM(16) = "전근무지총계"
    COLNAM(17) = "비과세계"
    COLNAM(18) = "총급여"
    COLNAM(19) = "근로소득공제"
    COLNAM(20) = "근로소득금액"
    COLNAM(21) = "본인공제금액"
    COLNAM(22) = "배우자유무"
    COLNAM(23) = "배우자공제액"
    COLNAM(24) = "부양가족수"
    COLNAM(25) = "부양가족공제"
    COLNAM(26) = "경로우대인원"
    COLNAM(27) = "경로우대공제"
    COLNAM(28) = "장애인인원"
    COLNAM(29) = "장애인공제액"
    COLNAM(30) = "부녀자유무"
    COLNAM(31) = "부녀자공제액"
    COLNAM(32) = "자녀양육인원"
    COLNAM(33) = "자녀양육공제"
    COLNAM(34) = "출산입양인원"
    COLNAM(35) = "출산입양공제"
    COLNAM(36) = "다자녀인원"
    COLNAM(37) = "다자녀공제"
    COLNAM(38) = "국민연금"
    COLNAM(39) = "기타연금(공무원연금)"
    COLNAM(40) = "기타연금(군인연금)"
    COLNAM(41) = "기타연금(사립학교교직원연금)"
    COLNAM(42) = "기타연금(별정우체국연금)"
    COLNAM(43) = "퇴직연금(근로자퇴직급여보장법)"
    COLNAM(44) = "퇴직연금(과학기술인공제)"
    COLNAM(45) = "보험료(건강보험)"
    COLNAM(46) = "보험료(고용보험)"
    COLNAM(47) = "보험료(보장성보험)"
    COLNAM(48) = "보험료(장애인전용)"
    COLNAM(49) = "의료비공제금액"
    COLNAM(50) = "교육비공제금액"
    COLNAM(51) = "주택임차차입금원리금상환-대출기관"
    COLNAM(52) = "주택임차차입금원리금상환-거주자"
    COLNAM(53) = "월세액"
    COLNAM(54) = "장기주택이자상환액-15년미만"
    COLNAM(55) = "장기주택이자상환액-29년이하"
    COLNAM(56) = "장기주택이자상환액-30년이상"
    COLNAM(57) = "기부금공제금액"
    COLNAM(58) = "혼인,이사,장례비"
    COLNAM(59) = "특별공제계"
    COLNAM(60) = "표준공제"
    COLNAM(61) = "차감소득금액"
    COLNAM(62) = "개인연금저축공제"
    COLNAM(63) = "연금저축소득공제"
    COLNAM(64) = "소기업공제부금소득공제"
    COLNAM(65) = "주택마련저축(청약저축)"
    COLNAM(66) = "주택마련저축(주택청약종합저축)"
    COLNAM(67) = "주택마련저축(장기주택마련저축)"
    COLNAM(68) = "주택마련저축(근로자주택마련저축)"
    COLNAM(69) = "투자조합출자공제"
    COLNAM(70) = "신용카드소득공제"
    COLNAM(71) = "우리사주조합공제"
    COLNAM(72) = "장기주식형저축소득공제"
    COLNAM(73) = "고용유지중소기업공제"
    COLNAM(74) = "기타소득공제계"
    COLNAM(75) = "종합소득과세표준"
    COLNAM(76) = "산출세액"
    COLNAM(77) = "소득법"
    COLNAM(78) = "조특법"
    COLNAM(79) = "조세조약"
    COLNAM(80) = "감면세액계"
    COLNAM(81) = "근로소득세액공제"
    COLNAM(82) = "납세조합공제"
    COLNAM(83) = "주택차입금"
    COLNAM(84) = "기부정처자금"
    COLNAM(85) = "외국납부"
    COLNAM(86) = "세액공제계"
    COLNAM(87) = "결정소득세"
    COLNAM(88) = "결정주민세"
    COLNAM(89) = "결정농특세"
    COLNAM(90) = "종(전)근무지_소득세"
    COLNAM(91) = "종(전)근무지_주민세"
    COLNAM(92) = "종(전)근무지_농특세"
    COLNAM(93) = "주(현)근무지_소득세"
    COLNAM(94) = "주(현)근무지_주민세"
    COLNAM(95) = "주(현)근무지_농특세"
    COLNAM(96) = "차감소득세"
    COLNAM(97) = "차감주민세"
    COLNAM(98) = "차감농특세"

    For iCol = 0 To 98
        oGrid1.Columns.Item(iCol).Editable = False
        oGrid1.Columns.Item(iCol).TitleObject.Caption = COLNAM(iCol)
        If iCol >= 8 Then
            oGrid1.Columns.Item(iCol).RightJustified = True
        End If
        If oCompany.version >= "860040" Then                       '2007B PL18 이상 일때(2007A 버전은 확인 필요)
             oGrid1.Columns.Item(iCol).TitleObject.Sortable = True
        End If

        
    Next iCol

    '// Link Button
    Set oEditCol = oGrid1.Columns.Item("EMPID")     '// 사원순번
    oEditCol.Type = gct_EditText
    oEditCol.LinkedObjectType = "171"
    
    Set oEditCol = oGrid1.Columns.Item("DOCNO1")    '// 월별자료No
    oEditCol.Type = gct_EditText
    oEditCol.LinkedObjectType = "171"
    Set oEditCol = oGrid1.Columns.Item("DOCNO2")    '// 소득항목No
    oEditCol.Type = gct_EditText
    oEditCol.LinkedObjectType = "171"
    Set oEditCol = oGrid1.Columns.Item("DOCNO3")    '// 종전문서No
    oEditCol.Type = gct_EditText
    oEditCol.LinkedObjectType = "171"
    Set oEditCol = oGrid1.Columns.Item("DOCNO4")    '// 세액계산No
    oEditCol.Type = gct_EditText
    oEditCol.LinkedObjectType = "171"
    Set oEditCol = oGrid1.Columns.Item("DOCNO5")    '// 의료비No
    oEditCol.Type = gct_EditText
    oEditCol.LinkedObjectType = "171"
    Set oEditCol = oGrid1.Columns.Item("DOCNO6")    '// 기부금No
    oEditCol.Type = gct_EditText
    oEditCol.LinkedObjectType = "171"
    Set oEditCol = oGrid1.Columns.Item("DOCNO7")    '// 연금.저축No
    oEditCol.Type = gct_EditText
    oEditCol.LinkedObjectType = "171"
    
    '// ComboBox
    Set oColumn = oGrid1.Columns.Item("JSNGBN")     '// 정산구분
    oColumn.Type = gct_ComboBox
    
    Set oComboCol = oGrid1.Columns.Item("JSNGBN")
    oComboCol.DisplayType = cdt_Description
    oComboCol.ValidValues.Add "1", "연말정산(재직자)"
    oComboCol.ValidValues.Add "2", "중도정산(퇴직자)"

    
    Set oColumn = oGrid1.Columns.Item("CLTCOD")     '// 사업장
    oColumn.Type = gct_ComboBox
    
    Set oComboCol = oGrid1.Columns.Item("CLTCOD")
    oComboCol.DisplayType = cdt_Description
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        Do Until oRecordSet.EOF
            oComboCol.ValidValues.Add oRecordSet.Fields(0).Value, oRecordSet.Fields(1).Value
            oRecordSet.MoveNext
        Loop
    End If
   
    oGrid1.AutoResizeColumns
    oForm.Freeze False

    Set oColumn = Nothing
    Set oEditCol = Nothing
    Set oComboCol = Nothing
    Set oRecordSet = Nothing
    Exit Sub
'/////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    oForm.Freeze False
    
    Set oColumn = Nothing
    Set oEditCol = Nothing
    Set oComboCol = Nothing
    Set oRecordSet = Nothing
    Sbo_Application.StatusBar.SetText "titleSetting 실행 중 오류가 발생했습니다." & Space$(10) & Err.Description, bmt_Short, smt_Error
End Sub

Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long)

    Select Case oUID
      Case "MSTCOD"
        If oForm.Items(oUID).Specific.String = "" Then
            oForm.DataSources.UserDataSources("MSTNAM").ValueEx = ""
        Else
            oForm.DataSources.UserDataSources("MSTNAM").ValueEx = MDC_SetMod.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & oForm.Items(oUID).Specific.String & "'", "")
        End If
    End Select
    oForm.Update
    
End Sub
'*******************************************************************
'// ItemEventHander
'*******************************************************************
Public Sub Raise_FormItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_FormItemEvent_Error
    
    Dim sQry        As String
    Dim i           As Long
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns     As SAPbouiCOM.Columns
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo Raise_FormItemEvent_Error
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Select Case pval.EventType
'et_ITEM_PRESSED''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case et_ITEM_PRESSED
        If pval.BeforeAction Then
            If pval.ItemUID = "Btn01" Then  '// 찾기 버튼
                Call Grid_Display
                BubbleEvent = False
            ElseIf pval.ItemUID = "CBtn01" And oForm.Items("MSTCOD").Enabled = True Then   '/ ChooseBtn 사원리스트
                oForm.Items("MSTCOD").CLICK ct_Regular
                Sbo_Application.ActivateMenuItem ("7425")
                BubbleEvent = False
            End If
        End If
        
'et_MATRIX_LINK_PRESSED'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case et_MATRIX_LINK_PRESSED
        If pval.BeforeAction Then
            If pval.ItemUID = "Grid1" And Left$(pval.ColUID, 5) = "DOCNO" Then
                Call UserFormLink(pval.ColUID, pval.Row)
                BubbleEvent = False
            End If
        End If

'et_VALIDATE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case et_VALIDATE
        If pval.BeforeAction = False And pval.ItemChanged = True Then
            If pval.ItemUID = "MSTCOD" Then
                Call FlushToItemValue(pval.ItemUID)
            End If
        End If

'et_FORM_UNLOAD'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case et_FORM_UNLOAD
        'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
        '컬렉션에서 삭제및 모든 메모리 제거
        'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
        If pval.BeforeAction = False Then
            RemoveForms oFormUniqueID
            Set oDS_ZPY507 = Nothing
            Set oGrid1 = Nothing
            Set oForm = Nothing
        End If
    End Select

    Exit Sub
'//////////////////////////////////////////////////////////////////////////////////////////////
Raise_FormItemEvent_Error:
    Sbo_Application.StatusBar.SetText "Raise_FormItemEvent_Error:" & Space$(10) & Err.Description, bmt_Short, smt_Error
End Sub

Private Sub UserFormLink(LinkName As String, LineNum As Long)
    Dim oTmpObject  As Variant
    Dim DocNum      As String
    Dim JSNYER      As String
    Dim CLTCOD      As String
    Dim MSTCOD      As String
    
    If Trim$(LinkName) <> "" Then
        Select Case LinkName
        Case "DOCNO1"
            Set oTmpObject = New ZPY343
            DocNum = oDS_ZPY507.GetValue("DOCNO1", LineNum)
            If Trim$(DocNum) <> "" Then
                Call oTmpObject.LoadForm(DocNum)
                Sbo_Application.Forms.ActiveForm.Select
            End If
        Case "DOCNO2"
            Set oTmpObject = New ZPY501
            JSNYER = Trim$(oDS_ZPY507.GetValue("JSNYMM", LineNum))
            CLTCOD = Trim$(oDS_ZPY507.GetValue("CLTCOD", LineNum))
            MSTCOD = Trim$(oDS_ZPY507.GetValue("MSTCOD", LineNum))
            If Trim$(JSNYER) <> "" And Trim$(CLTCOD) <> "" And Trim$(MSTCOD) <> "" Then
                Call oTmpObject.LoadForm(JSNYER, MSTCOD, CLTCOD)
                Sbo_Application.Forms.ActiveForm.Select
            End If
        Case "DOCNO3"
            Set oTmpObject = New ZPY502
            JSNYER = Trim$(oDS_ZPY507.GetValue("JSNYMM", LineNum))
            CLTCOD = Trim$(oDS_ZPY507.GetValue("CLTCOD", LineNum))
            MSTCOD = Trim$(oDS_ZPY507.GetValue("MSTCOD", LineNum))
            If Trim$(JSNYER) <> "" And Trim$(CLTCOD) <> "" And Trim$(MSTCOD) <> "" Then
                Call oTmpObject.LoadForm(JSNYER, MSTCOD, CLTCOD)
                Sbo_Application.Forms.ActiveForm.Select
            End If
        Case "DOCNO4"
            Set oTmpObject = New ZPY504
            DocNum = oDS_ZPY507.GetValue("DOCNO4", LineNum)
            If Trim$(DocNum) <> "" Then
                Call oTmpObject.LoadForm(DocNum)
                Sbo_Application.Forms.ActiveForm.Select
            End If
        Case "DOCNO5"
            Set oTmpObject = New ZPY506
            JSNYER = Trim$(oDS_ZPY507.GetValue("JSNYMM", LineNum))
            CLTCOD = Trim$(oDS_ZPY507.GetValue("CLTCOD", LineNum))
            MSTCOD = Trim$(oDS_ZPY507.GetValue("MSTCOD", LineNum))
            If Trim$(JSNYER) <> "" And Trim$(CLTCOD) <> "" And Trim$(MSTCOD) <> "" Then
                Call oTmpObject.LoadForm(JSNYER, MSTCOD, CLTCOD)
                Sbo_Application.Forms.ActiveForm.Select
            End If
        Case "DOCNO6"
            Set oTmpObject = New ZPY505
            JSNYER = Trim$(oDS_ZPY507.GetValue("JSNYMM", LineNum))
            CLTCOD = Trim$(oDS_ZPY507.GetValue("CLTCOD", LineNum))
            MSTCOD = Trim$(oDS_ZPY507.GetValue("MSTCOD", LineNum))
            If Trim$(JSNYER) <> "" And Trim$(CLTCOD) <> "" And Trim$(MSTCOD) <> "" Then
                Call oTmpObject.LoadForm(JSNYER, MSTCOD, CLTCOD)
                Sbo_Application.Forms.ActiveForm.Select
            End If
        Case "DOCNO7"
            Set oTmpObject = New ZPY508
            JSNYER = Trim$(oDS_ZPY507.GetValue("JSNYMM", LineNum))
            CLTCOD = Trim$(oDS_ZPY507.GetValue("CLTCOD", LineNum))
            MSTCOD = Trim$(oDS_ZPY507.GetValue("MSTCOD", LineNum))
            If Trim$(JSNYER) <> "" And Trim$(CLTCOD) <> "" And Trim$(MSTCOD) <> "" Then
                Call oTmpObject.LoadForm(JSNYER, MSTCOD, CLTCOD)
                Sbo_Application.Forms.ActiveForm.Select
            End If
        End Select
    
    End If
    Set oTmpObject = Nothing
End Sub

Private Sub Grid_Display()
On Error GoTo Error_Message
    Dim sQry            As String
    Dim ErrNum          As Integer
    Dim oRecordSet      As SAPbobsCOM.Recordset

    Dim FYEAR   As String:  Dim TYEAR   As String
    Dim JIGFDAT As String:  Dim JIGTDAT As String
    Dim SINFYMM As String:  Dim SINTYMM As String
    Dim MSTCOD  As String:  Dim MSTNAM  As String
    Dim CLTCOD  As String:  Dim JSNGBN  As String
    Dim PILMED  As Double:  Dim PILGBU  As Double
    Dim iRow    As Long
'/ Check
    ErrNum = 0
    iRow = 0
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    With oForm.DataSources.UserDataSources
        FYEAR = Trim$(.Item("FYEAR").ValueEx)
        TYEAR = Trim$(.Item("TYEAR").ValueEx)
        JIGFDAT = Trim$(.Item("JIGFDAT").ValueEx)
        JIGTDAT = Trim$(.Item("JIGTDAT").ValueEx)
        SINFYMM = Trim$(.Item("SINFYMM").ValueEx)
        SINTYMM = Trim$(.Item("SINTYMM").ValueEx)
        MSTCOD = Trim$(.Item("MSTCOD").ValueEx)
        MSTNAM = Trim$(.Item("MSTNAM").ValueEx)
        PILMED = Val(.Item("PILMED").ValueEx)
        PILGBU = Val(.Item("PILGBU").ValueEx)
    End With
    CLTCOD = oForm.Items("CLTCOD").Specific.Selected.Value
    JSNGBN = oForm.Items("JSNGBN").Specific.Selected.Value
    
    If FYEAR = "" Or TYEAR = "" Then        '// 정산년도 체크(입력 안할 경우 에러)
        ErrNum = 1
        GoTo Error_Message
    End If
    If JIGFDAT = "" Or JIGTDAT = "" Then    '// 지급일자 체크(입력 안할 경우 전체조회)
        JIGFDAT = "1900-01-01"
        JIGTDAT = "2999-12-31"
    End If
    If SINFYMM = "" Or SINTYMM = "" Then    '// 신고년월 체크(입력 안할 경우 전체조회)
        SINFYMM = "190001"
        SINTYMM = "299912"
    End If

    '/ 조회
    sQry = "Exec ZPY507 N'" & FYEAR & "', N'" & TYEAR & "', " & _
                       "N'" & JIGFDAT & "', N'" & JIGTDAT & "', " & _
                       "N'" & CLTCOD & "', N'" & MSTCOD & "', " & _
                       "N'" & MSTNAM & "', N'" & JSNGBN & "', " & _
                       "N'" & SINFYMM & "', N'" & SINTYMM & "', " & _
                        CStr(PILMED) & ", " & CStr(PILGBU)
    oDS_ZPY507.ExecuteQuery (sQry)
    iRow = oDS_ZPY507.Rows.Count
    If iRow = 1 Then
        oRecordSet.DoQuery sQry
        iRow = oRecordSet.RecordCount
    End If

    If iRow > 0 Then
        Sbo_Application.StatusBar.SetText iRow & " 건이 있습니다.", bmt_Short, smt_Success
    Else
        Sbo_Application.StatusBar.SetText "조회된 데이터가 없습니다.", bmt_Short, smt_Success
    End If
    Set oRecordSet = Nothing
    
    Call titleSetting
    Exit Sub
'/////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "기준년도는 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Success
    Else
        Sbo_Application.StatusBar.SetText "Grid_Display Error :" & Space$(10) & Err.Description, bmt_Short, smt_Error
    End If
End Sub

'*******************************************************************
'// MenuEventHander
'*******************************************************************
Public Sub Raise_FormMenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)

    If pval.BeforeAction = True Then
        Exit Sub
    End If

    Select Case pval.MenuUID
        Case "1287" '/ 복제
        Case "1281", "1282":
        Case "1288" To "1291":
        Case "1293"
    End Select
    Exit Sub
End Sub

Public Sub Raise_FormDataEvent(ByRef FormUID As String, ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
    Dim i       As Long
    Dim sQry    As String
    Dim oCombo  As SAPbouiCOM.ComboBox
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    
    On Error GoTo Raise_FormDataEvent_Error
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If (BusinessObjectInfo.BeforeAction = False) Then
        Select Case BusinessObjectInfo.EventType
            Case et_FORM_DATA_LOAD:     '//33
            Case et_FORM_DATA_ADD:      '//34
            Case et_FORM_DATA_UPDATE:   '//35
            Case et_FORM_DATA_DELETE:   '//36
        End Select

    End If
    Set oCombo = Nothing
    Set oRecordSet = Nothing
    Exit Sub
    
Raise_FormDataEvent_Error:
    Set oCombo = Nothing
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "Raise_FormDataEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True

End Sub

'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error
    Dim i           As Long
    Dim oXmlDoc             As New MSXML2.DOMDocument

    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\ZPY507.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = _
        oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)

    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '//여러개의 메트릭스가 틀경우에 층계모양처럼 로드 되도록 만든 모양
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetTotalFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetTotalFormsCount * 10)

    Sbo_Application.LoadBatchActions oXmlDoc.xml
    
    oFormUniqueID = "ZPY507_" & GetTotalFormsCount
    
    '폼 할당
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '컬렉션에 폼을 담는다   **컬렉션이란 개체를 담아 놓는 배열로서 여기서는 활성화되어져 있는 폼을 담고 있다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    AddForms Me, oFormUniqueID, "ZPY507"
    oForm.SupportedModes = -1

    '//////////////////////////////////////////////////////////////////////////////
    '***************************************************************
    '화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
    'oForm.DataBrowser.BrowseBy = "Code"
    '***************************************************************
    '//////////////////////////////////////////////////////////////////////////////
    oForm.Freeze True
    Call CreateItems

    oForm.EnableMenu ("1281"), False    '/ 추가
    oForm.EnableMenu ("1282"), False    '/ 추가

    oForm.Mode = fm_OK_MODE
    oForm.Freeze False
    oForm.Update
    oForm.Visible = True

    Set oXmlDoc = Nothing
    Exit Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
LoadForm_Error:
    Set oXmlDoc = Nothing
    Sbo_Application.StatusBar.SetText "Form_Load Error:" & Err.Description, bmt_Short, smt_Error
     If (oForm Is Nothing) = False Then
        oForm.Freeze False
        Set oForm = Nothing
    End If
End Sub

Private Sub CreateItems()
On Error GoTo Error_Message
    Dim oCombo          As SAPbouiCOM.ComboBox
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim oEdit           As SAPbouiCOM.EditText
    Dim sQry            As String
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    '/UserDataSource 선언
    With oForm.DataSources.UserDataSources
        .Add "FYEAR", dt_SHORT_TEXT, 4
        .Add "TYEAR", dt_SHORT_TEXT, 4
        .Add "JIGFDAT", dt_DATE
        .Add "JIGTDAT", dt_DATE
        .Add "MSTCOD", dt_SHORT_TEXT, 8
        .Add "MSTNAM", dt_LONG_TEXT, 200
        .Add "SINFYMM", dt_SHORT_TEXT, 6
        .Add "SINTYMM", dt_SHORT_TEXT, 6
        .Add "PILMED", dt_SUM
        .Add "PILGBU", dt_SUM
    End With
    
    Set oEdit = oForm.Items("FYEAR").Specific   '// 기준년도(From)
    oEdit.DataBind.SetBound True, "", "FYEAR"
    
    Set oEdit = oForm.Items("TYEAR").Specific   '// 기준년도(To)
    oEdit.DataBind.SetBound True, "", "TYEAR"
    
    Set oEdit = oForm.Items("JIGFDAT").Specific '// 지급일자(From)
    oEdit.DataBind.SetBound True, "", "JIGFDAT"
    
    Set oEdit = oForm.Items("JIGTDAT").Specific '// 지급일자(To)
    oEdit.DataBind.SetBound True, "", "JIGTDAT"
    
    Set oEdit = oForm.Items("MSTCOD").Specific  '// 사번
    oEdit.DataBind.SetBound True, "", "MSTCOD"
    
    Set oEdit = oForm.Items("MSTNAM").Specific  '// 성명
    oEdit.DataBind.SetBound True, "", "MSTNAM"
    
    Set oEdit = oForm.Items("SINFYMM").Specific '// 신고년월(From)
    oEdit.DataBind.SetBound True, "", "SINFYMM"
    
    Set oEdit = oForm.Items("SINTYMM").Specific '// 신고년월(To)
    oEdit.DataBind.SetBound True, "", "SINTYMM"
    
    Set oEdit = oForm.Items("PILMED").Specific  '// 의료비공제액
    oEdit.DataBind.SetBound True, "", "PILMED"
    
    Set oEdit = oForm.Items("PILGBU").Specific  '// 기부금공제액
    oEdit.DataBind.SetBound True, "", "PILGBU"
    

    '//사업장
    Set oCombo = oForm.Items("CLTCOD").Specific
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
    Call SetReDataCombo(oForm, sQry, oCombo)
    oCombo.ValidValues.Add "%", "전체"
    oForm.Items("CLTCOD").DisplayDesc = True
    oCombo.Select "%", psk_ByValue

    '// 정산구분
    Set oCombo = oForm.Items("JSNGBN").Specific
    oCombo.ValidValues.Add "%", "모두"
    oCombo.ValidValues.Add "1", "연말정산(재직자)"
    oCombo.ValidValues.Add "2", "중도정산(퇴직자)"
    oCombo.Select 0, psk_Index
    

    oForm.DataSources.UserDataSources("FYEAR").ValueEx = ZPAY_GBL_JSNYER
    oForm.DataSources.UserDataSources("TYEAR").ValueEx = ZPAY_GBL_JSNYER
    
'// 디비데이터 소스 개체 할당
    Set oGrid1 = oForm.Items("Grid1").Specific
    Set oDS_ZPY507 = oForm.DataSources.DataTables.Add("ZPY507")
    oDS_ZPY507.ExecuteQuery ("Exec ZPY507 '1900', '1900', NULL, NULL, '%', '%', '%', '%', '', '', 0, 0")
    oGrid1.DataTable = oDS_ZPY507

    Call titleSetting
    
    oForm.ActiveItem = "FYEAR"
    
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oRecordSet = Nothing
    Exit Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Error_Message:
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oRecordSet = Nothing
    
    Sbo_Application.StatusBar.SetText "CreateItems 실행 중 오류가 발생했습니다." & Space$(10) & Err.Description, bmt_Short, smt_Error
End Sub
