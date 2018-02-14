VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RPY505"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//  SAP MANAGE UI API 2004 SDK Sample
'//****************************************************************************
'//  File           : RPY505.cls
'//  Module         : 인사관리>정산관리>정산관련리포트
'//  Desc           : 소득자료제출집계표
'//  FormType       : 2010130505
'//  Create Date    : 2006.01.23
'//  Modified Date  : 2006.12.23
'//  Creator        : Ham Mi Kyoung
'//  Modifier       :
'//  Copyright  (c) Morning Data
'//****************************************************************************
Option Explicit

Public oFormUniqueID    As String
Public oForm            As SAPbouiCOM.Form

Private Sub Print_Query()
On Error GoTo Error_Message
    Dim WinTitle    As String
    Dim ReportName  As String
    Dim ErrNum      As Integer
    Dim sQry        As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Dim JSNFYMM     As String:  Dim JSNTYMM As String
    Dim SINFYMM     As String:  Dim SINTYMM As String
    Dim CLTCOD      As String
    Dim MSTDPT      As String
    Dim MSTCOD      As String
    Dim PRTDAT      As String
    Dim JOBGBN      As String
    Dim PRTTYP      As String
    
    '/ ODBC 연결 체크
    If ConnectODBC = False Then
        GoTo Error_Message
    End If
    
 '/ Default
    With oForm.DataSources.UserDataSources
        JSNFYMM = .Item("JSNFYMM").ValueEx
        JSNTYMM = .Item("JSNTYMM").ValueEx
        SINFYMM = .Item("SINFYMM").ValueEx
        SINTYMM = .Item("SINTYMM").ValueEx
        MSTCOD = .Item("MSTCOD").ValueEx
        PRTDAT = .Item("PRTDAT").ValueEx
    End With
    If Trim$(MSTCOD) = "" Then MSTCOD = "%"
    If Trim$(PRTDAT) = "" Then
        PRTDAT = Format$(Now, "yyyymmdd")
        oForm.DataSources.UserDataSources("PRTDAT").ValueEx = PRTDAT
    End If
    
    If MDC_SetMod.ChkYearMonth(JSNFYMM) = False Then JSNFYMM = ""
    If MDC_SetMod.ChkYearMonth(JSNTYMM) = False Then JSNTYMM = ""
    If MDC_SetMod.ChkYearMonth(SINFYMM) = False Then SINFYMM = ""
    If MDC_SetMod.ChkYearMonth(SINTYMM) = False Then SINTYMM = ""
    
 '/ Check
    ErrNum = 0
    Select Case True
    Case (JSNFYMM = "" Or JSNTYMM = "") And (SINFYMM = "" Or SINTYMM = "")
         ErrNum = 1
         GoTo Error_Message
    Case oForm.Items("Combo03").Specific.Selected Is Nothing
         ErrNum = 2
         GoTo Error_Message
    Case oForm.Items("Combo01").Specific.Selected Is Nothing
         ErrNum = 3
         GoTo Error_Message
    End Select
    CLTCOD = oForm.Items("Combo01").Specific.Selected.VALUE
    MSTDPT = oForm.Items("Combo02").Specific.Selected.VALUE
    JOBGBN = oForm.Items("Combo03").Specific.Selected.VALUE
    PRTTYP = oForm.Items("Combo04").Specific.Selected.VALUE  '/ 소득 구분
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    WinTitle = "소득자료제출집계표"
    ReportName = "RPY505.rpt"
    
    '/ Formula 수식필드***************************************************/
    ReDim gRpt_Formula(2)
    ReDim gRpt_Formula_Value(2)
    
'    gRpt_Formula(1) = "JSNYER":    gRpt_Formula_Value(1) = JSNYER
    gRpt_Formula(1) = "PRTDAT":    gRpt_Formula_Value(1) = PRTDAT
    
    If PRTTYP = "1" Then
        gRpt_Formula(2) = "PRTNAM": gRpt_Formula_Value(2) = "근로소득 (갑종)"
    ElseIf PRTTYP = "2" Then
        gRpt_Formula(2) = "PRTNAM": gRpt_Formula_Value(2) = "퇴직소득"
    ElseIf PRTTYP = "3" Then
        gRpt_Formula(2) = "PRTNAM": gRpt_Formula_Value(2) = "사업소득"
    ElseIf PRTTYP = "4" Then
        gRpt_Formula(2) = "PRTNAM": gRpt_Formula_Value(2) = "기타소득 (거주자)"
    ElseIf PRTTYP = "5" Then
        gRpt_Formula(2) = "PRTNAM": gRpt_Formula_Value(2) = "사업·기타소득 (비거주자)"
    ElseIf PRTTYP = "6" Then
        gRpt_Formula(2) = "PRTNAM": gRpt_Formula_Value(2) = "이자소득"
    ElseIf PRTTYP = "7" Then
        gRpt_Formula(2) = "PRTNAM": gRpt_Formula_Value(2) = "배당소득"
    End If
    
    WinTitle = "[RPY505] : " & WinTitle
 '/ SubReport /
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)
    
    gRpt_SRptSqry(1) = ""
    gRpt_SRptName(1) = ""
    
 '/ 조회조건문 /
    sQry = "Exec RPY505 " & "'" & Trim$(JSNFYMM) & "', '" & Trim$(JSNTYMM) & "', " & _
                            "'" & Trim$(SINFYMM) & "', '" & Trim$(SINTYMM) & "', " & _
                            "'" & Trim$(JOBGBN) & "', '" & Trim$(CLTCOD) & "', '" & Trim$(MSTDPT) & "', " & _
                            "'" & Trim$(MSTCOD) & "', '" & Trim$(PRTTYP) & "'"
 '/ Action /
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, 1, "Y", "V", "") = False Then
      '  SBO_Application.StatusBar.SetText "gCryReport_Action : 실패!", bmt_Short, smt_Error
    End If
    
    Set oRecordSet = Nothing
    Exit Sub
'/ Message /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
Error_Message:
    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "귀속 연월, 원천신고 연월이 모두 올바르지 않습니다.", bmt_Short, smt_Error
    ElseIf ErrNum = 2 Then
        Sbo_Application.StatusBar.SetText "출력 구분을 선택하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 3 Then
        Sbo_Application.StatusBar.SetText "자료 제출자를 선택하여 주십시오.", bmt_Short, smt_Error
    Else
        Sbo_Application.StatusBar.SetText "Print_Query : " & Space$(10) & Err.Description, bmt_Short, smt_Error
    End If
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
            If pval.ItemUID = "1" Then
                    If oForm.Mode = fm_OK_MODE Or oForm.Mode = fm_UPDATE_MODE Then
                        Call Print_Query
                        BubbleEvent = False
                    End If
            ElseIf pval.ItemUID = "CBtn1" Then
                If oForm.Items("MSTCOD").Enabled = True Then
                    oForm.Items("MSTCOD").CLICK ct_Regular
                    Sbo_Application.ActivateMenuItem ("7425")
                    BubbleEvent = False
                End If
            End If
        End If
'et_KEY_DOWN''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case et_KEY_DOWN
         If pval.BeforeAction = True And pval.ItemUID = "JsnYear" And pval.CharPressed = 9 Then
            If Len(Trim$(oForm.Items("JsnYear").Specific.String)) = 0 Then
                Sbo_Application.StatusBar.SetText "정산년도를 입력하여 주십시오.", bmt_Short, smt_Error
                BubbleEvent = False
            ElseIf Len(Trim$(oForm.Items(pval.ItemUID).Specific.String)) < 4 Then
                oForm.Items(pval.ItemUID).Specific.VALUE = Format$(oForm.Items(pval.ItemUID).Specific.VALUE, "2000")
            End If
         ElseIf pval.BeforeAction = True And pval.ItemUID = "MSTCOD" And pval.CharPressed = 9 Then
            If Trim$(oForm.Items("MSTCOD").Specific.String) <> "" Then
               If MDC_SetMod.Value_ChkYn("[@PH_PY001A]", "Code", "'" & Trim$(oForm.Items("MSTCOD").Specific.String) & "'") = True Then
                    Sbo_Application.StatusBar.SetText "사원번호를 확인하여 주십시오.", bmt_Short, smt_Error
                    BubbleEvent = False
               Else
                   oForm.Items("MSTNAM").Specific.VALUE = MDC_SetMod.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & Trim$(oForm.Items("MSTCOD").Specific.String) & "'", "")
               End If
            End If
        End If

'et_FORM_UNLOAD''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case et_FORM_UNLOAD
        'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
        '컬렉션에서 삭제및 모든 메모리 제거
        'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
        If pval.BeforeAction = False Then
           RemoveForms oFormUniqueID
            Set oForm = Nothing
        End If
    
    End Select
    
    Exit Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Raise_FormItemEvent_Error:
    Sbo_Application.StatusBar.SetText "Raise_FormItemEvent_Error:", bmt_Short, smt_Error
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

    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\RPY505.srf")
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
    
    oFormUniqueID = "RPY505_" & GetTotalFormsCount
    
    '폼 할당
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)
    
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '컬렉션에 폼을 담는다   **컬렉션이란 개체를 담아 놓는 배열로서 여기서는 활성화되어져 있는 폼을 담고 있다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    AddForms Me, oFormUniqueID, "RPY505"
    oForm.SupportedModes = -1
    oForm.Mode = fm_OK_MODE
    
    '//////////////////////////////////////////////////////////////////////////////
    '***************************************************************
    '화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
   ' oForm.DataBrowser.BrowseBy = "DocNum"
    '***************************************************************
    '//////////////////////////////////////////////////////////////////////////////
    oForm.Freeze True
    CreateItems
    oForm.Freeze False
    
    oForm.EnableMenu ("1281"), True  '/ 찾기
    oForm.EnableMenu ("1282"), False '/ 추가
    oForm.EnableMenu ("1284"), False '/ 취소
    oForm.EnableMenu ("1293"), False '/ 행삭제
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


'*******************************************************************
'
'*******************************************************************
Private Sub CreateItems()
On Error GoTo Error_Message
    Dim oCombo          As SAPbouiCOM.ComboBox
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim oEdit           As SAPbouiCOM.EditText
    Dim sQry            As String
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    With oForm.DataSources.UserDataSources
        .Add "JSNFYMM", dt_SHORT_TEXT, 6    '/ 귀속년월(From)
        .Add "JSNTYMM", dt_SHORT_TEXT, 6    '/ 귀속년월(To)
        .Add "SINFYMM", dt_SHORT_TEXT, 6    '/ 신고년월(From)
        .Add "SINTYMM", dt_SHORT_TEXT, 6    '/ 신고년월(To)
        .Add "MSTCOD", dt_SHORT_TEXT, 8
        .Add "MSTNAM", dt_SHORT_TEXT, 30
        .Add "PRTDAT", dt_DATE
    End With
    
    Set oEdit = oForm.Items("JSNFYMM").Specific
    oEdit.DataBind.SetBound True, "", "JSNFYMM"
    Set oEdit = oForm.Items("JSNTYMM").Specific
    oEdit.DataBind.SetBound True, "", "JSNTYMM"
    Set oEdit = oForm.Items("SINFYMM").Specific
    oEdit.DataBind.SetBound True, "", "SINFYMM"
    Set oEdit = oForm.Items("SINTYMM").Specific
    oEdit.DataBind.SetBound True, "", "SINTYMM"
    Set oEdit = oForm.Items("MSTCOD").Specific
    oEdit.DataBind.SetBound True, "", "MSTCOD"
    Set oEdit = oForm.Items("MSTNAM").Specific
    oEdit.DataBind.SetBound True, "", "MSTNAM"
    Set oEdit = oForm.Items("PRTDAT").Specific
    oEdit.DataBind.SetBound True, "", "PRTDAT"

'// Combo Box Setting
    '// 사업장
    Set oCombo = oForm.Items("Combo01").Specific
    oForm.Items("Combo01").DisplayDesc = True
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
    oRecordSet.DoQuery sQry
    Do Until oRecordSet.EOF
        oCombo.ValidValues.Add Trim$(oRecordSet.Fields(0).VALUE), Trim$(oRecordSet.Fields(1).VALUE)
        oRecordSet.MoveNext
    Loop
    If oCombo.ValidValues.Count > 0 Then
       Call oCombo.Select(0, psk_Index)
    End If
    
    '// 부서
    Set oCombo = oForm.Items("Combo02").Specific
    oForm.Items("Combo02").DisplayDesc = True
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '1' AND U_UseYN= 'Y'"
    oRecordSet.DoQuery sQry
    oCombo.ValidValues.Add "%", "모두"
    Do Until oRecordSet.EOF
        oCombo.ValidValues.Add Trim$(oRecordSet.Fields(0).VALUE), Trim$(oRecordSet.Fields(1).VALUE)
        oRecordSet.MoveNext
    Loop
    If oCombo.ValidValues.Count > 0 Then
       Call oCombo.Select(0, psk_Index)
    End If
    '// 생성구분
    Set oCombo = oForm.Items("Combo03").Specific
    oForm.Items("Combo03").DisplayDesc = True
    oCombo.ValidValues.Add "1", "연말정산(재직자)"
    oCombo.ValidValues.Add "2", "중도정산(퇴직자)"
    oCombo.ValidValues.Add "3", "전체"
    Call oCombo.Select("3", psk_ByValue)
'// 소득종류
    Set oCombo = oForm.Items("Combo04").Specific
    oForm.Items("Combo04").DisplayDesc = True
    oCombo.ValidValues.Add "1", "근로소득(갑종)"
    oCombo.ValidValues.Add "2", "퇴직소득"
    oCombo.ValidValues.Add "3", "사업소득(거주자)"
    oCombo.ValidValues.Add "4", "기타소득(거주자)"
    oCombo.ValidValues.Add "5", "사업.기타소득(비거주자)"
    oCombo.ValidValues.Add "6", "이자소득"
    oCombo.ValidValues.Add "7", "배당소득"
    
    If oCombo.ValidValues.Count > 0 Then
       Call oCombo.Select(0, psk_Index)
    End If
'/
    oForm.DataSources.UserDataSources("JSNFYMM").ValueEx = Format$(DateAdd("YYYY", -1, Now), "YYYY") & "01"
    oForm.DataSources.UserDataSources("JSNTYMM").ValueEx = Format$(DateAdd("YYYY", -1, Now), "YYYY") & "12"
    oForm.DataSources.UserDataSources("SINFYMM").ValueEx = Format$(DateAdd("YYYY", -1, Now), "YYYY") & "01"
    oForm.DataSources.UserDataSources("SINTYMM").ValueEx = Format$(Now, "YYYYMM")
    
    
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oRecordSet = Nothing
   Exit Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Error_Message:
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oRecordSet = Nothing
    Sbo_Application.StatusBar.SetText "CreateItems Error:" & Err.Description, bmt_Short, smt_Error
End Sub
