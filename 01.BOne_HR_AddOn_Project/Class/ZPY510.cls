VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ZPY510"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//  SAP MANAGE UI API 2004 SDK Sample
'//****************************************************************************
'//  File           : ZPY510.cls
'//  Module         : 원천징수>근로소득
'//  Desc           : 종전근무지 일괄생성
'//  FormType       : 2010110510
'//  Create Date    : 2010.01.05
'//  Modified Date  :
'//  Creator        : Choi Dong Kwon
'//  Copyright  (c) Morning Data
'//****************************************************************************
Option Explicit
Public oFormUniqueID    As String
Public oForm            As SAPbouiCOM.Form
Private oGrid           As SAPbouiCOM.Grid
Private oDS_ZPY510      As SAPbouiCOM.DataTable
Private mJSNYER         As String

'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error
    Dim i           As Long
    Dim oXmlDoc             As New MSXML2.DOMDocument

    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\ZPY510.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = _
        oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    
    '------------------------------------------------------------------------
    '//여러개의 메트릭스가 틀경우에 층계모양처럼 로드 되도록 만든 모양
    '------------------------------------------------------------------------
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetTotalFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetTotalFormsCount * 10)

    Sbo_Application.LoadBatchActions oXmlDoc.xml
    
    oFormUniqueID = "ZPY510_" & GetTotalFormsCount
    
    '폼 할당
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)
    
    '--------------------------------------------------------------------------------------------------------------
    '컬렉션에 폼을 담는다   **컬렉션이란 개체를 담아 놓는 배열로서 여기서는 활성화되어져 있는 폼을 담고 있다
    '--------------------------------------------------------------------------------------------------------------
    
    AddForms Me, oFormUniqueID, "ZPY510"
    oForm.SupportedModes = -1
    oForm.Mode = fm_OK_MODE
    
    '//////////////////////////////////////////////////////////////////////////////
    '***************************************************************
    '화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
    'oForm.DataBrowser.BrowseBy = "DocNum"
    '***************************************************************
    '//////////////////////////////////////////////////////////////////////////////
    
    Call CreateItems
    
    oForm.EnableMenu ("1293"), False  '/ 행삭제
    oForm.EnableMenu ("1284"), False  '/ 취소

    oForm.Update
    oForm.Visible = True
    
    Set oXmlDoc = Nothing
    Exit Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
LoadForm_Error:
    Set oXmlDoc = Nothing
     Sbo_Application.StatusBar.SetText "종전근무지 일괄생성을 실행시킬 수 없습니다.", bmt_Short, smt_Error
     If (oForm Is Nothing) = False Then
        Set oForm = Nothing
    End If
End Sub

'*******************************************************************
' Item Initial
'*******************************************************************
Private Sub CreateItems()
On Error GoTo Error_Message
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim oCombo          As SAPbouiCOM.ComboBox
    Dim oEdit           As SAPbouiCOM.EditText
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '// UserDataSources
    With oForm.DataSources.UserDataSources
        .Add "JSNYER", dt_SHORT_TEXT, 4
        .Add "MSTCOD", dt_SHORT_TEXT, 8
        .Add "MSTNAM", dt_SHORT_TEXT, 30
    End With
    
    Set oEdit = oForm.Items("JSNYER").Specific
    oEdit.DataBind.SetBound True, "", "JSNYER"
    Set oEdit = oForm.Items("MSTCOD").Specific
    oEdit.DataBind.SetBound True, "", "MSTCOD"
    Set oEdit = oForm.Items("MSTNAM").Specific
    oEdit.DataBind.SetBound True, "", "MSTNAM"
    
    oForm.DataSources.UserDataSources("JSNYER").ValueEx = Format$(Now, "YYYY")
    
    '//사업장
    Set oCombo = oForm.Items("FCLTCOD").Specific
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
    Call SetReDataCombo(oForm, sQry, oCombo)
    oCombo.ValidValues.Add "%", "전체"
    oForm.Items("FCLTCOD").DisplayDesc = True
    oCombo.Select "%", psk_ByValue
    
    Set oCombo = oForm.Items("TCLTCOD").Specific
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
    Call SetReDataCombo(oForm, sQry, oCombo)
    oCombo.ValidValues.Add "%", "전체"
    oForm.Items("TCLTCOD").DisplayDesc = True
    oCombo.Select "%", psk_ByValue
    
    '/ Grid
    Set oGrid = oForm.Items("Grid1").Specific
    Call oForm.DataSources.DataTables.Add("ZPY510")
    
    Set oDS_ZPY510 = oForm.DataSources.DataTables.Item("ZPY510")
    Call oDS_ZPY510.ExecuteQuery("EXEC ZPY510 '1900', '', '', ''")
    oGrid.DataTable = oDS_ZPY510
    
    Call titleSetting
        
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oRecordSet = Nothing
    Exit Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Error_Message:
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oRecordSet = Nothing
    Sbo_Application.StatusBar.SetText "CreateItems Error :" & Space$(10) & Err.Description, bmt_Short, smt_Error
End Sub

'---------------------------------------------------------------------------------------
' Procedure : TitleSetting
' Author    : Choi Dong Kwon
' Date      : 2008-07-15
' Purpose   : Grid의 Column Title 지정
'---------------------------------------------------------------------------------------
'
Private Sub titleSetting()
On Error GoTo Error_Message
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim ErrNum          As Integer
    Dim i               As Integer
    
    Dim COLNAM(24)      As String

'/ Initial
    ErrNum = 0
    
    oForm.Freeze True
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
'/  컬럼명
    COLNAM(0) = "선택"
    COLNAM(1) = "사원번호"
    COLNAM(2) = "성명"
    COLNAM(3) = "사원순번"
    COLNAM(4) = "자사코드"
    COLNAM(5) = "종전자사코드"
    COLNAM(6) = "사업자번호"
    COLNAM(7) = "귀속시작일"
    COLNAM(8) = "귀속종료일"
    COLNAM(9) = "당해감면시작일"
    COLNAM(10) = "당해감면종료일"
    COLNAM(11) = "급여금액"
    COLNAM(12) = "상여금액"
    COLNAM(13) = "인정상여"
    COLNAM(14) = "주식매수선택권행사이익"
    COLNAM(15) = "우리사주조합인출"
    COLNAM(16) = "비과세총계"
    COLNAM(17) = "건강보험"
    COLNAM(18) = "고용보험"
    COLNAM(19) = "국민연금"
    COLNAM(20) = "연금보험료"
    COLNAM(21) = "소득세"
    COLNAM(22) = "주민세"
    COLNAM(23) = "농특세"
    COLNAM(24) = "퇴직연금"
    
    '// 컬럼명 셋팅
    For i = 0 To 24
        oGrid.Columns.Item(i).TitleObject.Caption = COLNAM(i)
        
        If i >= 11 Then
            oGrid.Columns.Item(i).RightJustified = True
        Else
            oGrid.Columns.Item(i).RightJustified = False
        End If
        
        If i > 0 Then
            oGrid.Columns.Item(i).Editable = False
        Else
            oGrid.Columns.Item(i).Editable = True
        End If
        
    Next i
    
    '// Grid의 컬럼별 ComboBox, CheckBox 세팅
    Call Grid_Col_Define
    oGrid.AutoResizeColumns

'/
    oForm.Freeze False

    Set oRecordSet = Nothing
    Exit Sub
'/////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    oForm.Freeze False
    Set oRecordSet = Nothing
    
    If ErrNum = 1 Then
    
    Else
        Sbo_Application.StatusBar.SetText "titleSetting 실행 중 오류가 발생했습니다." & Space$(10) & Err.Description, bmt_Short, smt_Error
    End If
End Sub

Private Sub Grid_Display()
On Error GoTo Error_Message
    Dim sQry        As String
    Dim ErrNum      As Integer
    Dim oRecordSet  As SAPbobsCOM.Recordset
    Dim iRow        As Integer
    
    Dim MSTCOD      As String
    Dim FCLTCOD     As String
    Dim TCLTCOD     As String
    
'/  Default Value
    ErrNum = 0
    iRow = 0

    mJSNYER = Trim$(oForm.DataSources.UserDataSources("JSNYER").ValueEx)
    MSTCOD = Trim$(oForm.DataSources.UserDataSources("MSTCOD").ValueEx)

'/ Check
    Select Case True
    Case oForm.Items("FCLTCOD").Specific.Selected Is Nothing
        ErrNum = 1
        GoTo Error_Message
    Case oForm.Items("TCLTCOD").Specific.Selected Is Nothing
        ErrNum = 1
        GoTo Error_Message
    Case mJSNYER = ""
        ErrNum = 2
        GoTo Error_Message
    End Select
    FCLTCOD = oForm.Items("FCLTCOD").Specific.Selected.Value
    TCLTCOD = oForm.Items("TCLTCOD").Specific.Selected.Value
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    oForm.Mode = fm_OK_MODE
'/ 조회
    sQry = " Exec ZPY510 '" & Trim$(mJSNYER) & "', '" & Trim$(MSTCOD) & "', '" & _
                              Trim$(FCLTCOD) & "', '" & Trim$(TCLTCOD) & "'"
    Debug.Print sQry
    Call oDS_ZPY510.ExecuteQuery(sQry)
    iRow = oForm.DataSources.DataTables.Item(0).Rows.Count
    
    If iRow = 1 Then
        oRecordSet.DoQuery sQry
        iRow = oRecordSet.RecordCount
    End If

    Sbo_Application.StatusBar.SetText iRow & " 건이 있습니다.", bmt_Short, smt_Success
    Set oRecordSet = Nothing
    Call titleSetting
    
    Exit Sub
'/////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "자사 코드는 필수입니다. 선택하여 주십시오.", bmt_Short, smt_Success
    ElseIf ErrNum = 2 Then
        Sbo_Application.StatusBar.SetText "귀속년도는 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Success
    Else
        Sbo_Application.StatusBar.SetText "Grid_Display Error :" & Space$(10) & Err.Description, bmt_Short, smt_Error
    End If
End Sub

Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long)

    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    oForm.Freeze True
    
    '// 사번 입력시 성명 조회
    If oUID = "MSTCOD" Then
        If oForm.Items(oUID).Specific.Value = "" Then
            oForm.DataSources.UserDataSources.Item("MSTNAM").ValueEx = ""
        Else
            oForm.DataSources.UserDataSources.Item("MSTNAM").ValueEx = _
            MDC_SetMod.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", oForm.Items(oUID).Specific.Value)
        End If
        
        oForm.Update
    End If
    oForm.Freeze False
End Sub

'*******************************************************************
'// ItemEventHander
'*******************************************************************
Public Sub Raise_FormItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
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
                If pval.ItemUID = "Btn1" Then  '// 폴더 열기 버튼
                    Call Grid_Display
                    BubbleEvent = False
                ElseIf pval.ItemUID = "Btn2" Then  '// File Upload 버튼
                    Call Grid_Save
                    BubbleEvent = False
                ElseIf pval.ItemUID = "CBtn1" And oForm.Items("MSTCOD").Enabled = True Then   '/ ChooseBtn사원리스트
                    oForm.Items("MSTCOD").CLICK ct_Regular
                    Sbo_Application.ActivateMenuItem ("7425")
                    BubbleEvent = False
                ElseIf pval.ItemUID = "Grid1" And pval.ColUID = "U_CHECK" And pval.Row = -1 Then                   '/ 전체체크
                    Call CheckAll
                End If
            End If

'et_KEY_DOWN''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_KEY_DOWN
            If pval.BeforeAction = False Then
                If pval.ItemUID = "Grid1" And pval.ColUID = "U_MSTCOD" And pval.CharPressed = 9 Then
                    If MDC_SetMod.Value_ChkYn("[@PH_PY001A]", "Code", "'" + oDS_ZPY510.GetValue("U_MSTCOD", pval.Row) + "'") = True Or _
                        oDS_ZPY510.GetValue("U_MSTCOD", pval.Row) = "" Then
                        oGrid.Columns.Item("U_MSTCOD").CLICK pval.Row
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                End If
            End If
            
'et_VALIDATE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_VALIDATE
            If pval.BeforeAction = False And pval.ItemChanged = True Then
                If pval.ItemUID = "MSTCOD" Then
                    FlushToItemValue pval.ItemUID
                ElseIf pval.ItemUID = "Grid1" Then
                    Call FlushToItemValue(pval.ColUID, pval.Row)
                End If
            End If

'et_FORM_UNLOAD''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_FORM_UNLOAD
            '----------------------------------------------------
            '컬렉션에서 삭제및 모든 메모리 제거
            '----------------------------------------------------
            If pval.BeforeAction = False Then
                RemoveForms oFormUniqueID
                Set oDS_ZPY510 = Nothing
                Set oGrid = Nothing
                Set oForm = Nothing
            End If
    End Select

    Exit Sub
'//////////////////////////////////////////////////////////////////////////////////////////////
Raise_FormItemEvent_Error:
    Sbo_Application.StatusBar.SetText "Raise_FormItemEvent_Error:" & Space$(10) & Err.Description, bmt_Short, smt_Error
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

Private Sub CheckAll()
    Dim CheckType   As String
    Dim oRow        As Long

    oForm.Freeze True
    CheckType = "Y"
    For oRow = 0 To oGrid.Rows.Count - 1
        If Trim$(oDS_ZPY510.GetValue("U_CHECK", oRow)) = "N" Then
            CheckType = "N"
            Exit For
        End If
    Next
    
    For oRow = 0 To oGrid.Rows.Count - 1
        oDS_ZPY510.Rows.Offset = oRow
        If CheckType = "N" Then
            oDS_ZPY510.setValue "U_CHECK", oRow, "Y"
        Else
            oDS_ZPY510.setValue "U_CHECK", oRow, "N"
        End If
    Next
    oForm.Freeze False
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Grid_Col_Define
' Author    : Choi Dong Kwon
' Date      : 2008-07-14
' Purpose   : Grid의 Column들에 대하여 LinkButton, ComboBox, CheckBox등을 정의
'---------------------------------------------------------------------------------------
'
Private Sub Grid_Col_Define()

On Error GoTo Error_Message

    Dim sQry            As String
    Dim ErrNum          As Integer

    Dim oColumn         As SAPbouiCOM.GridColumn
    Dim oEditCol        As SAPbouiCOM.EditTextColumn
    Dim oComboCol       As SAPbouiCOM.ComboBoxColumn
    Dim oRecordSet      As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Set oColumn = oGrid.Columns.Item("U_CHECK")
    oColumn.Type = gct_CheckBox
    
    '// 사원순번에 LinkButton 추가
    Set oEditCol = oGrid.Columns.Item("EMPID")
    oEditCol.Type = gct_EditText
    oEditCol.LinkedObjectType = "171"
    
    '// 자사코드
    '// EditText Column => ComboBox Column으로 변경
    Set oColumn = oGrid.Columns.Item("CLTCOD")
    oColumn.Type = gct_ComboBox
    
    Set oComboCol = oGrid.Columns.Item("CLTCOD")
    oComboCol.DisplayType = cdt_Description
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
    oRecordSet.DoQuery sQry
    Do Until oRecordSet.EOF
        oComboCol.ValidValues.Add oRecordSet.Fields(0).Value, oRecordSet.Fields(1).Value
        oRecordSet.MoveNext
    Loop
    
    Set oColumn = oGrid.Columns.Item("JCLTCOD")
    oColumn.Type = gct_ComboBox
    
    Set oComboCol = oGrid.Columns.Item("JCLTCOD")
    oComboCol.DisplayType = cdt_Description
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
    oRecordSet.DoQuery sQry
    Do Until oRecordSet.EOF
        oComboCol.ValidValues.Add oRecordSet.Fields(0).Value, oRecordSet.Fields(1).Value
        oRecordSet.MoveNext
    Loop
    
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
    If ErrNum = 1 Then
    
    Else
        Sbo_Application.StatusBar.SetText "Grid_Col_Define 실행 중 오류가 발생했습니다." & Space$(10) & Err.Description, bmt_Short, smt_Error
    End If

End Sub

'---------------------------------------------------------------------------------------
' Procedure : Grid_Save
' Author    : Choi Dong Kwon
' Date      : 2008-07-15
' Purpose   : Grid의 내용을 일괄 저장하는 프로시저
'---------------------------------------------------------------------------------------
'
Private Sub Grid_Save()

On Error GoTo Error_Message

    Dim oRecordSet  As SAPbobsCOM.Recordset
    Dim sQry        As String
    Dim ErrNum      As Integer
    
    Dim oRow        As Long
    Dim UserId      As Long
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    UserId = oCompany.UserSignature
    
    ErrNum = 0
    For oRow = 0 To oGrid.Rows.Count - 1
    
        '// 체크된 행만 저장
        If oDS_ZPY510.GetValue("U_CHECK", oRow) = "Y" Then
        
            oCompany.StartTransaction       '/ 트랜잭션 시작
            
            sQry = "EXEC ZPY510_1 '" & mJSNYER & "', " & _
                                 "'" & oDS_ZPY510.GetValue("MSTCOD", oRow) & "', " & _
                                 "'" & oDS_ZPY510.GetValue("CLTCOD", oRow) & "', " & _
                                 "'" & oDS_ZPY510.GetValue("JCLTCOD", oRow) & "', " & _
                                       CStr(UserId) & " "
            oRecordSet.DoQuery sQry
            Debug.Print sQry
            oCompany.EndTransaction wf_Commit   '/ 트랜잭션 종료
        End If
    Next
    
    Call Grid_Display
    Sbo_Application.StatusBar.SetText "종전근무지 일괄생성이 완료되었습니다.", bmt_Short, smt_Success
    'oForm.Mode = fm_OK_MODE
    Set oRecordSet = Nothing
    Exit Sub
    
Error_Message:
    
    oCompany.EndTransaction wf_RollBack '/ 트랜잭션 RollBack
    
    Set oRecordSet = Nothing
    Sbo_Application.StatusBar.SetText "Grid_Save 실행 중 오류가 발생했습니다." & Space$(10) & Err.Description, bmt_Short, smt_Error

End Sub
