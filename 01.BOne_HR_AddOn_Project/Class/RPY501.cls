VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RPY501"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//  SAP MANAGE UI API 2004 SDK Sample
'//****************************************************************************
'//  File           : RPY501.cls
'//  Module         : 인사관리>정산관리>정산관련리포트
'//  Desc           : 월별 자료 현황
'//  FormType       : 2010130501
'//  Create Date    : 2006.01.10
'//  Modified Date  : 2006.12.10
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
    
    Dim JSNYER      As String
    Dim STRMON      As String
    Dim ENDMON      As String
    Dim JOBGBN      As String
    Dim CLTCOD      As String
    Dim Branch      As String
    Dim MSTDPT      As String
    Dim MSTCOD      As String
    Dim PRTGBN      As String
    
    '/ ODBC 연결 체크
    If ConnectODBC = False Then
        GoTo Error_Message
    End If
    
 '/ Default
    JSNYER = oForm.Items("JsnYear").Specific.String
    STRMON = Format$(oForm.Items("SMonth").Specific.String, "00")
    ENDMON = Format$(oForm.Items("EMonth").Specific.String, "00")
    MSTCOD = oForm.Items("MSTCOD").Specific.String
    If Trim$(MSTCOD) = "" Then MSTCOD = "%"
 '/ Check
    ErrNum = 0
    Select Case True
    Case MDC_SetMod.ChkYearMonth(JSNYER & STRMON) = False, _
         MDC_SetMod.ChkYearMonth(JSNYER & ENDMON) = False
         ErrNum = 1
         GoTo Error_Message
    Case oForm.Items("Combo03").Specific.Selected Is Nothing
         ErrNum = 2
         GoTo Error_Message
    Case oForm.Items("CLTCOD").Specific.Selected Is Nothing
         ErrNum = 3
         GoTo Error_Message
    End Select
    CLTCOD = oForm.Items("CLTCOD").Specific.Selected.VALUE
    
    MSTDPT = oForm.Items("Combo02").Specific.Selected.VALUE
    JOBGBN = oForm.Items("Combo03").Specific.Selected.VALUE
    PRTGBN = oForm.Items("PRTGBN").Specific.Selected.VALUE
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    If PRTGBN = "1" Then
        WinTitle = "월별 자료 현황(집계표)"
        ReportName = "RPY501.rpt"
    Else
        WinTitle = "월별 자료 현황(집계표)"
        ReportName = "RPY501_1.rpt"
    End If
        
    '/ Formula 수식필드***************************************************/
    ReDim gRpt_Formula(2)
    ReDim gRpt_Formula_Value(2)
    
    gRpt_Formula(1) = "CLTNAM":    gRpt_Formula_Value(1) = oCompany.CompanyName
    gRpt_Formula(2) = "PRTLMT":    gRpt_Formula_Value(2) = Mid$(JSNYER, 1, 4) & "년 "
    
    WinTitle = "[RPY501] : " & WinTitle
 '/ SubReport /
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
   
    gRpt_SRptSqry(1) = ""
    gRpt_SRptName(1) = ""
    
 '/ 조회조건문 /
    If PRTGBN = "1" Then
        sQry = "Exec RPY501 " & "'" & Trim$(JSNYER) & "', '" & Trim$(STRMON) & "', " & _
                                "'" & Trim$(ENDMON) & "', '" & Trim$(JOBGBN) & "', " & _
                                "'" & Trim$(CLTCOD) & "', " & _
                                "'" & Trim$(MSTDPT) & "', '" & Trim$(MSTCOD) & "'"
    Else
        sQry = "Exec RPY501_1 " & "'" & Trim$(JSNYER) & "', '" & Trim$(STRMON) & "', " & _
                                  "'" & Trim$(ENDMON) & "', '" & Trim$(JOBGBN) & "', " & _
                                  "'" & Trim$(CLTCOD) & "', " & _
                                  "'" & Trim$(MSTDPT) & "', '" & Trim$(MSTCOD) & "'"
    End If
 
 '/ Action /
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, 1, "Y", "V", "", 2) = False Then
      '  SBO_Application.StatusBar.SetText "gCryReport_Action : 실패!", bmt_Short, smt_Error
    End If
    
    Set oRecordSet = Nothing
    Exit Sub
'/ Message /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
Error_Message:
    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "기준 연월을 확인하세요.", bmt_Short, smt_Error
    ElseIf ErrNum = 2 Then
        Sbo_Application.StatusBar.SetText "출력 구분을 선택 하세요..", bmt_Short, smt_Error
    ElseIf ErrNum = 3 Then
        Sbo_Application.StatusBar.SetText "자사 코드를 선택 하세요..", bmt_Short, smt_Error
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
'et_GOT_FOCUS''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_GOT_FOCUS

'et_FORM_UNLOAD''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_FORM_UNLOAD
            '----------------------------------------------------
            '컬렉션에서 삭제및 모든 메모리 제거
            '----------------------------------------------------
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
'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm()

    Dim i                   As Long
    Dim oXmlDoc             As New MSXML2.DOMDocument

    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\RPY501.srf")
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
    
    oFormUniqueID = "RPY501_" & GetTotalFormsCount
    
    '폼 할당
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)
    
    AddForms Me, oFormUniqueID, "RPY501"
    oForm.SupportedModes = -1
    oForm.Mode = fm_OK_MODE
    
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
        .Add "JsnYear", dt_SHORT_TEXT, 4    '/ 생성년도
        .Add "SMonth", dt_SHORT_TEXT, 2      '/ 시작월
        .Add "EMonth", dt_SHORT_TEXT, 2      '/ 종료월
        .Add "MSTCOD", dt_SHORT_TEXT, 8
        .Add "MSTNAM", dt_SHORT_TEXT, 30
    End With
    
    Set oEdit = oForm.Items("JsnYear").Specific
    oEdit.DataBind.SetBound True, "", "JsnYear"
    Set oEdit = oForm.Items("SMonth").Specific
    oEdit.DataBind.SetBound True, "", "SMonth"
    Set oEdit = oForm.Items("EMonth").Specific
    oEdit.DataBind.SetBound True, "", "EMonth"
    Set oEdit = oForm.Items("MSTCOD").Specific
    oEdit.DataBind.SetBound True, "", "MSTCOD"
    Set oEdit = oForm.Items("MSTNAM").Specific
    oEdit.DataBind.SetBound True, "", "MSTNAM"

    '// 자사코드
    Set oCombo = oForm.Items("CLTCOD").Specific
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
    oRecordSet.DoQuery sQry
    oCombo.ValidValues.Add "%", "모두"
    Do Until oRecordSet.EOF
        oCombo.ValidValues.Add Trim$(oRecordSet.Fields(0).VALUE), Trim$(oRecordSet.Fields(1).VALUE)
        oRecordSet.MoveNext
    Loop
    If oCombo.ValidValues.Count > 0 Then
        oCombo.Select 0, psk_Index
    End If
    
'    '// 지점
'    Set oCombo = oForm.Items("Combo01").Specific
'    oForm.Items("Combo01").DisplayDesc = True
'    sQry = "SELECT Code, Name FROM OUBR WHERE Code <> '-2' OR (Code = '-2' AND Name <> N'주요') ORDER BY Code ASC"
'    oRecordSet.DoQuery sQry
'    oCombo.ValidValues.Add "%", "모두"
'    Do Until oRecordSet.EOF
'        oCombo.ValidValues.Add Trim$(oRecordSet.Fields(0).Value), Trim$(oRecordSet.Fields(1).Value)
'        oRecordSet.MoveNext
'    Loop
'    If oCombo.ValidValues.Count > 0 Then
'       Call oCombo.Select(0, psk_Index)
'    End If
    
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
    
    
    '// 출력구분
    Set oCombo = oForm.Items("PRTGBN").Specific
    oForm.Items("PRTGBN").DisplayDesc = True
    oCombo.ValidValues.Add "1", "집계표"
    oCombo.ValidValues.Add "2", "명세서"
    oCombo.Select 0, psk_Index
    
    '// Initial Value
    oForm.DataSources.UserDataSources("JsnYear").ValueEx = Year(Now)
    oForm.DataSources.UserDataSources("SMonth").ValueEx = "01"
    oForm.DataSources.UserDataSources("EMonth").ValueEx = "12"
    
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
