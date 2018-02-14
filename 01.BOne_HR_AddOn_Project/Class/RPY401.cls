VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RPY401"
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
    
    Dim STRDAT   As String * 10
    Dim ENDDAT   As String * 10
    Dim PRTDAT   As String
    Dim PRTGBN   As String * 1
    Dim JOBGBN   As String
    Dim Branch   As String
    Dim MSTDPT   As String
    Dim MSTCOD   As String
    Dim STRJIG   As String * 10
    Dim ENDJIG   As String * 10
    
 '/ Default
    STRDAT = Mid$(oForm.Items("STRDAT").Specific.String, "1", "4") & "-" & Mid$(oForm.Items("STRDAT").Specific.String, "6", "2") & "-" & Mid$(oForm.Items("STRDAT").Specific.String, "9", "2")
    ENDDAT = Mid$(oForm.Items("ENDDAT").Specific.String, "1", "4") & "-" & Mid$(oForm.Items("ENDDAT").Specific.String, "6", "2") & "-" & Mid$(oForm.Items("ENDDAT").Specific.String, "9", "2")
    If Trim$(oForm.Items("PRTDAT").Specific.String) = "" Then
        oForm.Items("PRTDAT").Specific.VALUE = Format$(Now, "yyyymmdd")
    End If
    PRTDAT = oForm.Items("PRTDAT").Specific.VALUE
    PRTDAT = Mid$(PRTDAT, 1, 4) & "년  " & Mid$(PRTDAT, 5, 2) & "월  " & Mid$(PRTDAT, 7, 2) & "일"
    JOBGBN = oForm.Items("Combo03").Specific.Selected.VALUE
    Branch = IIf(oForm.Items("Combo01").Specific.Selected.VALUE = "-1", "%", oForm.Items("Combo01").Specific.Selected.VALUE)
    MSTDPT = IIf(oForm.Items("Combo02").Specific.Selected.VALUE = "-1", "%", oForm.Items("Combo02").Specific.Selected.VALUE)
    MSTCOD = IIf(oForm.Items.Item("MSTCOD").Specific.VALUE = "", "%", oForm.Items.Item("MSTCOD").Specific.VALUE)
    STRJIG = Mid$(oForm.Items("STRJIG").Specific.String, "1", "4") & "-" & Mid$(oForm.Items("STRJIG").Specific.String, "6", "2") & "-" & Mid$(oForm.Items("STRJIG").Specific.String, "9", "2")
    ENDJIG = Mid$(oForm.Items("ENDJIG").Specific.String, "1", "4") & "-" & Mid$(oForm.Items("ENDJIG").Specific.String, "6", "2") & "-" & Mid$(oForm.Items("ENDJIG").Specific.String, "9", "2")
    
 '/ Check
    ErrNum = 0
    Select Case True
    Case IsDate(STRDAT) = False, IsDate(ENDDAT) = False
        ErrNum = 1
        GoTo Error_Message
    Case STRDAT > ENDDAT
        ErrNum = 2
        GoTo Error_Message
    Case oForm.Items("Combo03").Specific.Selected Is Nothing
        ErrNum = 3
        GoTo Error_Message
    Case STRDAT >= "2010-01-01"
        Set oRecordSet = Nothing
        Exit Sub
    Case IsDate(STRJIG) = False, IsDate(ENDJIG) = False
        ErrNum = 5
        GoTo Error_Message
    Case STRJIG > ENDJIG
        ErrNum = 6
        GoTo Error_Message
    End Select
    PRTGBN = oForm.DataSources.UserDataSources("OptionDS").ValueEx
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    WinTitle = "퇴직소득원천영수증"
    
    '/ Formula 수식필드***************************************************/
    ReDim ZRpt_Formula(2)
    ReDim ZRpt_Formula_Value(2)
    
    ZRpt_Formula(1) = "PRTDAT":    ZRpt_Formula_Value(1) = PRTDAT
    ZRpt_Formula(2) = "PRTGBN":    ZRpt_Formula_Value(2) = PRTGBN
    
    WinTitle = "[RPY401] : " & WinTitle
    
    ReportName = "RPY401.rpt"
 '/ SubReport /
    ReDim ZRpt_SRptSqry(1)
    ReDim ZRpt_SRptName(1)
   
    ZRpt_SRptSqry(1) = ""
    ZRpt_SRptName(1) = ""
 '/ ParameterFields
 '/ 조회조건문 /
    sQry = "Exec RPY401 " & "'" & Trim$(STRDAT) & "', '" & Trim$(ENDDAT) & "', '" & _
                                  Trim$(JOBGBN) & "', '" & Trim$(Branch) & "', '" & _
                                  Trim$(MSTDPT) & "', '" & Trim$(MSTCOD) & "', '" & _
                                  Trim$(STRJIG) & "', '" & Trim$(ENDJIG) & "'"
                                  
'    oRecordSet.DoQuery sQry
'    If oRecordSet.RecordCount = 0 Then
'        ErrNum = 4
'        GoTo error_Message
'    End If
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
        Sbo_Application.StatusBar.SetText "대상 기간을 확인하세요.", bmt_Short, smt_Error
    ElseIf ErrNum = 2 Then
        Sbo_Application.StatusBar.SetText "대상종료일자가 시작일자보다 작습니다. 확인하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 3 Then
        Sbo_Application.StatusBar.SetText "출력구분을 선택 하세요.", bmt_Short, smt_Error
    ElseIf ErrNum = 4 Then
        Sbo_Application.StatusBar.SetText "해당하는 자료가 존재하지 않습니다.", bmt_Short, smt_Error
    ElseIf ErrNum = 5 Then
        Sbo_Application.StatusBar.SetText "지급 기간을 확인하세요.", bmt_Short, smt_Error
    ElseIf ErrNum = 6 Then
        Sbo_Application.StatusBar.SetText "지급종료일자가 시작일자보다 작습니다. 확인하여 주십시오.", bmt_Short, smt_Error
    Else
        Sbo_Application.StatusBar.SetText "Print_Query : " & Space$(10) & Err.Description, bmt_Short, smt_Error
    End If
End Sub
Private Sub Print_Query2()
On Error GoTo Error_Message
    Dim WinTitle    As String
    Dim ReportName  As String
    Dim ErrNum      As Integer
    Dim sQry        As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Dim STRDAT   As String * 10
    Dim ENDDAT   As String * 10
    Dim PRTDAT   As String
    Dim PRTGBN   As String * 1
    Dim JOBGBN   As String
    Dim Branch   As String
    Dim MSTDPT   As String
    Dim MSTCOD   As String
    Dim STRJIG   As String * 10
    Dim ENDJIG   As String * 10
    
 '/ Default
    STRDAT = Mid$(oForm.Items("STRDAT").Specific.String, "1", "4") & "-" & Mid$(oForm.Items("STRDAT").Specific.String, "6", "2") & "-" & Mid$(oForm.Items("STRDAT").Specific.String, "9", "2")
    ENDDAT = Mid$(oForm.Items("ENDDAT").Specific.String, "1", "4") & "-" & Mid$(oForm.Items("ENDDAT").Specific.String, "6", "2") & "-" & Mid$(oForm.Items("ENDDAT").Specific.String, "9", "2")
    If Trim$(oForm.Items("PRTDAT").Specific.String) = "" Then
        oForm.Items("PRTDAT").Specific.VALUE = Format$(Now, "yyyymmdd")
    End If
    PRTDAT = oForm.Items("PRTDAT").Specific.VALUE
    PRTDAT = Mid$(PRTDAT, 1, 4) & "년  " & Mid$(PRTDAT, 5, 2) & "월  " & Mid$(PRTDAT, 7, 2) & "일"
    JOBGBN = oForm.Items("Combo03").Specific.Selected.VALUE
    Branch = IIf(oForm.Items("Combo01").Specific.Selected.VALUE = "-1", "%", oForm.Items("Combo01").Specific.Selected.VALUE)
    MSTDPT = IIf(oForm.Items("Combo02").Specific.Selected.VALUE = "-1", "%", oForm.Items("Combo02").Specific.Selected.VALUE)
    MSTCOD = IIf(oForm.Items.Item("MSTCOD").Specific.VALUE = "", "%", oForm.Items.Item("MSTCOD").Specific.VALUE)
    STRJIG = Mid$(oForm.Items("STRJIG").Specific.String, "1", "4") & "-" & Mid$(oForm.Items("STRJIG").Specific.String, "6", "2") & "-" & Mid$(oForm.Items("STRJIG").Specific.String, "9", "2")
    ENDJIG = Mid$(oForm.Items("ENDJIG").Specific.String, "1", "4") & "-" & Mid$(oForm.Items("ENDJIG").Specific.String, "6", "2") & "-" & Mid$(oForm.Items("ENDJIG").Specific.String, "9", "2")
    
 '/ Check
    ErrNum = 0
    Select Case True
    Case IsDate(STRDAT) = False, IsDate(ENDDAT) = False
        ErrNum = 1
        GoTo Error_Message
    Case STRDAT > ENDDAT
        ErrNum = 2
        GoTo Error_Message
    Case oForm.Items("Combo03").Specific.Selected Is Nothing
        ErrNum = 3
        GoTo Error_Message
    Case IsDate(STRJIG) = False, IsDate(ENDJIG) = False
        ErrNum = 5
        GoTo Error_Message
    Case STRJIG > ENDJIG
        ErrNum = 6
        GoTo Error_Message
    End Select
    PRTGBN = oForm.DataSources.UserDataSources("OptionDS").ValueEx
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    WinTitle = "퇴직소득원천징수영수증"
    
    '/ Formula 수식필드***************************************************/
    ReDim ZRpt_Formula(2)
    ReDim ZRpt_Formula_Value(2)
    
    ZRpt_Formula(1) = "PRTDAT":    ZRpt_Formula_Value(1) = PRTDAT
    ZRpt_Formula(2) = "PRTGBN":    ZRpt_Formula_Value(2) = PRTGBN
    
    WinTitle = "[RPY401_2] : " & WinTitle
    If ENDDAT <= "2009-12-31" Then
        ReportName = "RPY401_2.rpt"
    ElseIf ENDDAT <= "2012-07-26" Then
        ReportName = "RPY401_2010.rpt"
    Else
        ReportName = "RPY401_2012.rpt"
    End If
 '/ SubReport /
    ReDim ZRpt_SRptSqry(1)
    ReDim ZRpt_SRptName(1)
   
    ZRpt_SRptSqry(1) = ""
    ZRpt_SRptName(1) = ""
 '/ ParameterFields
 '/ 조회조건문 /
    sQry = "Exec RPY401_2 " & "'" & Trim$(STRDAT) & "', '" & Trim$(ENDDAT) & "', '" & _
                                    Trim$(JOBGBN) & "', '" & Trim$(Branch) & "', '" & _
                                    Trim$(MSTDPT) & "', '" & Trim$(MSTCOD) & "', '" & _
                                    Trim$(STRJIG) & "', '" & Trim$(ENDJIG) & "'"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
        Set oRecordSet = Nothing
        Exit Sub
    End If


 '/ Action /
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, 2, "Y", "V", "") = False Then
      '  SBO_Application.StatusBar.SetText "gCryReport_Action : 실패!", bmt_Short, smt_Error
    End If
    
    Set oRecordSet = Nothing
    Exit Sub
'/ Message /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
Error_Message:
    Set oRecordSet = Nothing
    
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "대상 기간을 확인하세요.", bmt_Short, smt_Error
    ElseIf ErrNum = 2 Then
        Sbo_Application.StatusBar.SetText "대상종료일자가 시작일자보다 작습니다. 확인하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 3 Then
        Sbo_Application.StatusBar.SetText "출력구분을 선택 하세요.", bmt_Short, smt_Error
    ElseIf ErrNum = 4 Then
        Sbo_Application.StatusBar.SetText "해당하는 자료가 존재하지 않습니다.", bmt_Short, smt_Error
    ElseIf ErrNum = 5 Then
        Sbo_Application.StatusBar.SetText "지급 기간을 확인하세요.", bmt_Short, smt_Error
    ElseIf ErrNum = 6 Then
        Sbo_Application.StatusBar.SetText "지급종료일자가 시작일자보다 작습니다. 확인하여 주십시오.", bmt_Short, smt_Error
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
'                        If oForm.Mode = fm_OK_MODE Or oForm.Mode = fm_UPDATE_MODE Then
                            Call Print_Query
                            Call Print_Query2
                            BubbleEvent = False
'                        End If
                ElseIf pval.ItemUID = "CBtn1" Then
                    If oForm.Items("MSTCOD").Enabled = True Then
                        oForm.Items("MSTCOD").CLICK ct_Regular
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                End If
            End If
'et_VALIDATE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_VALIDATE
            If pval.BeforeAction = False And pval.ItemChanged = True And _
              (pval.ItemUID = "STRDAT" Or pval.ItemUID = "ENDDAT" Or _
              pval.ItemUID = "MSTCOD") Then
               FlushToItemValue pval.ItemUID
            End If
'et_KEY_DOWN''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_KEY_DOWN
             If pval.BeforeAction = True And pval.ItemUID = "MSTCOD" And pval.CharPressed = 9 Then
                If Trim$(oForm.Items("MSTCOD").Specific.String) <> "" Then
                   If MDC_SetMod.Value_ChkYn("OHEM", "U_MSTCOD", "'" & Trim$(oForm.Items("MSTCOD").Specific.String) & "'") = True Then
                        Sbo_Application.StatusBar.SetText "사원번호를 확인하여 주십시오.", bmt_Short, smt_Error
                        BubbleEvent = False
                   Else
                       oForm.Items("MSTNAM").Specific.VALUE = MDC_SetMod.Get_ReData("LastName+FirstName", "U_MSTCOD", "OHEM", "'" & Trim$(oForm.Items("MSTCOD").Specific.String) & "'", "")
                   End If
                End If
            End If
'et_GOT_FOCUS''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_GOT_FOCUS

'et_FORM_UNLOAD''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_FORM_UNLOAD
            'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
            '컬렉션에서 삭제및 모든 메모리 제거
            'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
            If pval.BeforeAction = False Then
               RemoveForms oForm.uniqueID
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
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\RPY401.srf")
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
    
    oFormUniqueID = "RPY401_" & GetTotalFormsCount
    
    '폼 할당
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)
    
    AddForms Me, oFormUniqueID, "RPY401"
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

      Dim oCombo          As SAPbouiCOM.ComboBox
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim oEdit           As SAPbouiCOM.EditText
    Dim sQry            As String
    Dim oOption         As SAPbouiCOM.OptionBtn
       
    On Error GoTo Error_Message
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Call oForm.DataSources.UserDataSources.Add("STRDAT", dt_DATE, 10)      '/ 시작월
    Call oForm.DataSources.UserDataSources.Add("ENDDAT", dt_DATE, 10)      '/ 종료월
    Call oForm.DataSources.UserDataSources.Add("MSTCOD", dt_SHORT_TEXT, 8)
    Call oForm.DataSources.UserDataSources.Add("MSTNAM", dt_SHORT_TEXT, 30)
    Call oForm.DataSources.UserDataSources.Add("PRTDAT", dt_DATE, 10)
    Call oForm.DataSources.UserDataSources.Add("STRJIG", dt_DATE, 10)      '/ 지급시작일
    Call oForm.DataSources.UserDataSources.Add("ENDJIG", dt_DATE, 10)      '/ 지급종료월
    
    Set oEdit = oForm.Items("STRDAT").Specific
    oEdit.DataBind.SetBound True, "", "STRDAT"
    Set oEdit = oForm.Items("ENDDAT").Specific
    oEdit.DataBind.SetBound True, "", "ENDDAT"
    Set oEdit = oForm.Items("MSTCOD").Specific
    oEdit.DataBind.SetBound True, "", "MSTCOD"
    Set oEdit = oForm.Items("MSTNAM").Specific
    oEdit.DataBind.SetBound True, "", "MSTNAM"
    Set oEdit = oForm.Items("PRTDAT").Specific
    oEdit.DataBind.SetBound True, "", "PRTDAT"
    Set oEdit = oForm.Items("STRJIG").Specific
    oEdit.DataBind.SetBound True, "", "STRJIG"
    Set oEdit = oForm.Items("ENDJIG").Specific
    oEdit.DataBind.SetBound True, "", "ENDJIG"
    
'// Combo Box Setting
    '// 지점
    Set oCombo = oForm.Items("Combo01").Specific
    oForm.Items("Combo01").DisplayDesc = True
    
    sQry = "SELECT Code, Name FROM [@PH_PY005A] "
        oRecordSet.DoQuery sQry
        
        If oRecordSet.RecordCount > 0 Then
            Do Until oRecordSet.EOF
                oCombo.ValidValues.Add oRecordSet.Fields(0).VALUE, oRecordSet.Fields(1).VALUE
                oRecordSet.MoveNext
            Loop
        Else
            oCombo.ValidValues.Add "", "-"
        End If
        
    Call CLTCOD_Select(oForm, "Combo01")
'    sQry = "SELECT Code, Name FROM OUBR WHERE Code <> '-2' OR (Code = '-2' AND Name <> N'주요') ORDER BY Code ASC"
'    oRecordSet.DoQuery sQry
'    oCombo.ValidValues.Add "%", "모두"
'    Do Until oRecordSet.EOF
'        oCombo.ValidValues.Add Trim$(oRecordSet.Fields(0).VALUE), Trim$(oRecordSet.Fields(1).VALUE)
'        oRecordSet.MoveNext
'    Loop
'    If oCombo.ValidValues.Count > 0 Then
'       Call oCombo.Select(0, psk_Index)
'    End If
    '// 부서
    Set oCombo = oForm.Items("Combo02").Specific
    oForm.Items("Combo02").DisplayDesc = True
     sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '1' AND U_UseYN= 'Y' AND U_Char2 = '" & oForm.Items.Item("Combo01").Specific.VALUE & "'"
    Call SetReDataCombo(oForm, sQry, oCombo, "Y")
    
'    sQry = " SELECT U_MSTDPT, Name FROM OUDP  WHERE ISNULL(U_MSTDPT, '') <> '' ORDER BY U_MSTDPT ASC "
'    oRecordSet.DoQuery sQry
'    oCombo.ValidValues.Add "%", "모두"
'    Do Until oRecordSet.EOF
'        oCombo.ValidValues.Add Trim$(oRecordSet.Fields(0).VALUE), Trim$(oRecordSet.Fields(1).VALUE)
'        oRecordSet.MoveNext
'    Loop
'    If oCombo.ValidValues.Count > 0 Then
'       Call oCombo.Select(0, psk_Index)
'    End If
    '// 생성구분
    Set oCombo = oForm.Items("Combo03").Specific
    oForm.Items("Combo03").DisplayDesc = True
    oCombo.ValidValues.Add "%", "모두"
    oCombo.ValidValues.Add "1", "퇴직정산"
    oCombo.ValidValues.Add "2", "중도정산"
    Call oCombo.Select(0, psk_Index)   '/ 전체
    
  
     '//옵션버튼(생성방법)
    Call oForm.DataSources.UserDataSources.Add("OptionDS", dt_SHORT_TEXT, 1)

    oForm.Items("Opt1").Visible = True
    oForm.Items("Opt2").Visible = True
    oForm.Items("Opt3").Visible = True
    
    Set oOption = oForm.Items("Opt1").Specific
    oOption.DataBind.SetBound True, "", "OptionDS"
    oOption.ValOn = "1": oOption.ValOff = "N"
    
    Set oOption = oForm.Items("Opt2").Specific
    oOption.DataBind.SetBound True, "", "OptionDS"
    oOption.GroupWith ("Opt1")
    If oOption.ValOn <> "2" Then
       oOption.ValOn = "2"
    End If
    oOption.ValOff = "N"
    
    Set oOption = oForm.Items("Opt3").Specific
    oOption.DataBind.SetBound True, "", "OptionDS"
    oOption.GroupWith ("Opt1")
    If oOption.ValOn <> "3" Then
       oOption.ValOn = "3"
    End If
    oOption.ValOff = "N"
    
    Set oOption = oForm.Items("Opt1").Specific
    oOption.Selected = True

    oForm.DataSources.UserDataSources("STRDAT").ValueEx = Format$(Now, "YYYY") & "0101"
    oForm.DataSources.UserDataSources("ENDDAT").ValueEx = Format$(Now, "YYYYMMDD")
    
    
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

Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long)
        
    Select Case oUID
      Case "STRDAT"
            ZPAY_GBL_JSNYER = Left$(oForm.Items(oUID).Specific.VALUE, 4)
            oForm.DataSources.UserDataSources("ENDDAT").ValueEx = ZPAY_GBL_JSNYER & "1231"
            oForm.Items("ENDDAT").Update
      Case "MSTCOD"
            If oForm.Items(oUID).Specific.String = "" Then
                oForm.DataSources.UserDataSources("MSTNAM").ValueEx = ""
            Else
                oForm.DataSources.UserDataSources("MSTNAM").ValueEx = MDC_SetMod.Get_ReData("LastName+FirstName", "U_MSTCOD", "OHEM", "'" & Trim$(oForm.Items("MSTCOD").Specific.String) & "'", "")
            End If
            oForm.Items("MSTNAM").Update
    End Select
    oForm.Items(oUID).Update
End Sub

