VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ZPY341"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//  SAP MANAGE UI API 2004 SDK Sample
'//****************************************************************************
'//  File           : ZPY341.cls
'//  Module         : 인사관리>정산관리
'//  Desc           : 월별 정산자료 생성
'//  FormType       : 2010110341
'//  Create Date    : 2005.12.12
'//  Modified Date  :
'//  Creator        : Ham Mi Kyoung
'//  Modifier       :
'//  Copyright  (c) Morning Data
'//****************************************************************************
Option Explicit

Public oFormUniqueID    As String
Public oForm            As SAPbouiCOM.Form

Private oMat1           As SAPbouiCOM.Matrix
Private Last_Item       As String                     '클래스에서 선택한 마지막 아이템 Uid값

Private oJsnYear        As String
Private oSMonth        As String
Private oEMonth        As String
Private oJsnGbn        As String

Private Type TmpZPY341
    JOBDAT As String * 6
    CLTCOD As String * 8
    MSTCOD As String * 10
    MSTNAM As String * 30
    JIGBIL As String * 8
    
    'TAMT(1 To 23)  As Double
    TAMT(1 To 45)  As Double '2009연말정산항목추가(2009.12.31)
End Type
Dim tmpArr As TmpZPY341

'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error
    Dim i           As Long
    Dim oXmlDoc             As New MSXML2.DOMDocument

    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\ZPY341.srf")
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
    
    oFormUniqueID = "ZPY341_" & GetTotalFormsCount
    
    '폼 할당
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)
    
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '컬렉션에 폼을 담는다   **컬렉션이란 개체를 담아 놓는 배열로서 여기서는 활성화되어져 있는 폼을 담고 있다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    AddForms Me, oFormUniqueID, "ZPY341"
    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
    
    oForm.Freeze True
    Call CreateItems
    Call SetDocument(oFromDocEntry01)
    oForm.Freeze False
    
    oForm.EnableMenu ("1281"), False '/ 찾기
    oForm.EnableMenu ("1282"), True  '/ 추가
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
                If pval.ItemUID = "CBtn1" Then   '/ ChooseBtn사원리스트
                    oForm.Items("MstCode").CLICK ct_Regular
                    Sbo_Application.ActivateMenuItem ("7425")
                    BubbleEvent = False
                ElseIf pval.ItemUID = "Opt1" Or pval.ItemUID = "Opt2" Then
                    If oForm.DataSources.UserDataSources("OptionDS").ValueEx = "1" Then
                        oForm.Items("MstCode").CLICK ct_Regular
                        oForm.Items("Path").Enabled = False
                        oForm.Items("Btn1").Enabled = False
                    Else
                        oForm.Items("Path").Enabled = True
                        oForm.Items("Btn1").Enabled = True
                    End If
                ElseIf pval.ItemUID = "Btn1" And oForm.Mode = fm_ADD_MODE Then  '// 경로선택
                    Call Excell_Upload
                    BubbleEvent = False
                ElseIf pval.ItemUID = "Execution" Then   '/ 월별자료생성실행
                    If Execution_Process = False Then
                        BubbleEvent = False
                    Else
                        oForm.ActiveItem = "JsnYear"
                        oForm.Items("JsnGbn").Enabled = False
                        oForm.Items("STRDAT").Enabled = False
                        oForm.Items("ENDDAT").Enabled = False
                        oForm.Items("CLTCOD").Enabled = False
'                        oForm.Items("BPLId").Enabled = False
                        oForm.Items("DptStr").Enabled = False
                        oForm.Items("DptEnd").Enabled = False
                        oForm.Items("MstCode").Enabled = False
                    End If
                    If oForm.Mode <> fm_ADD_MODE Then
                        oForm.Mode = fm_ADD_MODE
                    End If
                ElseIf pval.ItemUID = "1" And oForm.Mode = fm_ADD_MODE Then
                    If Execution_Save = False Then
                        BubbleEvent = False
                        Exit Sub
                    Else
                        oForm.ActiveItem = "JsnYear"
                        oForm.Items("JsnGbn").Enabled = True
                        oForm.Items("STRDAT").Enabled = True
                        oForm.Items("ENDDAT").Enabled = True
                        oForm.Items("CLTCOD").Enabled = True
'                        oForm.Items("BPLId").Enabled = True
                        oForm.Items("DptStr").Enabled = True
                        oForm.Items("DptEnd").Enabled = True
                        oForm.Items("MstCode").Enabled = True
                    End If
                End If
            Else
                If pval.ItemUID = "1" And pval.ActionSuccess = True And oForm.Mode = fm_OK_MODE Then
                    FormItemEnabled
                End If
            End If
'et_VALIDATE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_VALIDATE
            If pval.BeforeAction = False And pval.ItemChanged = True And _
              (pval.ItemUID = "JsnYear" Or pval.ItemUID = "SMonth" Or pval.ItemUID = "EMonth" Or _
              pval.ItemUID = "MstCode") Then
               FlushToItemValue pval.ItemUID
            End If
'et_CLICK''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_CLICK
             If pval.BeforeAction = True And pval.ItemUID <> "1000001" And pval.ItemUID <> "2" Then
                If Last_Item = "JsnYear" Then     '/정산년도
                    If Trim$(oForm.Items(Last_Item).Specific.VALUE) <> "" Then
                        If MDC_SetMod.ChkYearMonth(Trim$(CStr((oForm.Items(Last_Item).Specific.VALUE))) & "01") = False Then
                            oForm.Items(Last_Item).Update
                            Sbo_Application.StatusBar.SetText "정산년도를 확인하여 주십시오.", bmt_Short, smt_Error
                            BubbleEvent = False
                        End If
                    End If
                ElseIf Last_Item = "SMonth" Or Last_Item = "EMonth" Then
                    If Trim$(oForm.Items(Last_Item).Specific.VALUE) <> "" Then
                        If MDC_SetMod.ChkYearMonth(oJsnYear & Format$(oForm.Items(Last_Item).Specific.VALUE, "00")) = False Then
                            oForm.Items(Last_Item).Update
                            Sbo_Application.StatusBar.SetText "생성기간을 확인하여 주십시오.", bmt_Short, smt_Error
                            BubbleEvent = False
                        End If
                    End If
                ElseIf Last_Item = "MstCode" Then
                    If Trim$(oForm.Items(Last_Item).Specific.String) <> "" And _
                    MDC_SetMod.Value_ChkYn("[@PH_PY001A]", "Code", "'" & Trim$(oForm.Items(Last_Item).Specific.String) & "'", "") = True Then
                        oForm.Items(Last_Item).Update
                        Sbo_Application.StatusBar.SetText "사원번호를 확인하여 주십시오.", bmt_Short, smt_Error
                        BubbleEvent = False
                    End If
                End If
            End If
'et_KEY_DOWN''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_KEY_DOWN
             If pval.BeforeAction = True And pval.ItemUID = "JsnYear" And pval.CharPressed = 9 Then
                If Len(Trim$(oForm.Items(pval.ItemUID).Specific.String)) < 4 Then
                    oForm.Items(pval.ItemUID).Specific.VALUE = Format$(oForm.Items(pval.ItemUID).Specific.VALUE, "2000")
                End If
                If MDC_SetMod.ChkYearMonth(Trim$(CStr((oForm.Items(pval.ItemUID).Specific.VALUE))) & "01") = False Then
                    Sbo_Application.StatusBar.SetText "정산년도를 확인하여 주십시오.", bmt_Short, smt_Error
                    BubbleEvent = False
                End If
             ElseIf pval.BeforeAction = True And (pval.ItemUID = "SMonth" Or pval.ItemUID = "EMonth") And pval.CharPressed = 9 Then
                If MDC_SetMod.ChkYearMonth(oJsnYear & Format$(oForm.Items(pval.ItemUID).Specific.VALUE, "00")) = False Then
                    Sbo_Application.StatusBar.SetText "생성기간을 확인하여 주십시오.", bmt_Short, smt_Error
                    BubbleEvent = False
                End If
             ElseIf pval.BeforeAction = True And pval.ItemUID = "MstCode" And pval.CharPressed = 9 Then
                If Trim$(oForm.Items("MstCode").Specific.String) <> "" And _
                    MDC_SetMod.Value_ChkYn("[@PH_PY001A]", "Code", "'" & Trim$(oForm.Items("MstCode").Specific.String) & "'", "") = True Then
                    Sbo_Application.StatusBar.SetText "사원번호를 확인하여 주십시오.", bmt_Short, smt_Error
                    BubbleEvent = False
                End If
            End If
'et_COMBO_SELECT''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_COMBO_SELECT
            If pval.BeforeAction = False And pval.ItemChanged = True Then
                If pval.ItemUID = "JsnGbn" Then
                    FlushToItemValue pval.ItemUID
                End If
                If pval.ItemUID = "CLTCOD" Then
                    '//기본사항 - 부서1 (사업장에 따른 부서변경)
                    Set oCombo = oForm.Items("DptStr").Specific

                    If oCombo.ValidValues.Count > 0 Then
                        For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                            oCombo.ValidValues.Remove i, psk_Index
                        Next i
                        oCombo.ValidValues.Add "", ""
                        oCombo.Select 0, psk_Index
                    End If
                    
                    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
                    sQry = sQry & " WHERE Code = '1' AND U_Char2 = '" & Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) & "'"
                    sQry = sQry & " ORDER BY U_Code"
                    Call SetReDataCombo(oForm, sQry, oCombo)
                    oCombo.ValidValues.Add "%", "전체"
                    oCombo.Select "%", psk_ByValue
                    
                    '//기본사항 - 부서2 (사업장에 따른 부서변경)
                    Set oCombo = oForm.Items("DptEnd").Specific

                    If oCombo.ValidValues.Count > 0 Then
                        For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                            oCombo.ValidValues.Remove i, psk_Index
                        Next i
                        oCombo.ValidValues.Add "", ""
                        oCombo.Select 0, psk_Index
                    End If
                    
                    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
                    sQry = sQry & " WHERE Code = '1' AND U_Char2 = '" & Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) & "'"
                    sQry = sQry & " ORDER BY U_Code"
                    Call SetReDataCombo(oForm, sQry, oCombo)
                    oCombo.ValidValues.Add "%", "전체"
                    oCombo.Select "%", psk_ByValue
                End If
            End If
'et_GOT_FOCUS''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_GOT_FOCUS
            If Last_Item = "Mat1" Then
                If pval.Row > 0 Then
                    Last_Item = pval.ItemUID
                End If
            Else
                Last_Item = pval.ItemUID
            End If
'et_FORM_UNLOAD''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_FORM_UNLOAD
            'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
            '컬렉션에서 삭제및 모든 메모리 제거
            'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
            If pval.BeforeAction = True Then
                Call TmpTable_Delete
            End If
            If pval.BeforeAction = False Then
               RemoveForms oFormUniqueID
                Set oForm = Nothing
                Set oMat1 = Nothing
            End If
    End Select
    
    Exit Sub
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////
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
            FormItemEnabled
            oForm.Items("JsnYear").CLICK ct_Regular
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

Private Sub SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY001_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call FormItemEnabled
    Else
        oForm.Mode = fm_FIND_MODE
        Call FormItemEnabled
        oForm.Items("Code").Specific.VALUE = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY001_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub CreateItems()
On Error GoTo Error_Message
    Dim oCombo1         As SAPbouiCOM.ComboBox
    Dim oCombo2         As SAPbouiCOM.ComboBox
    Dim oOption         As SAPbouiCOM.OptionBtn
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim oEdit           As SAPbouiCOM.EditText
    Dim oColumn         As SAPbouiCOM.Column
    Dim sQry            As String

    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
        
    Call oForm.DataSources.UserDataSources.Add("JsnYear", dt_SHORT_TEXT, 4)    '/ 생성년도
    Call oForm.DataSources.UserDataSources.Add("JsnGbn", dt_SHORT_TEXT, 10)     '/ 생성구분
    Call oForm.DataSources.UserDataSources.Add("CLTCOD", dt_SHORT_TEXT, 10)      '/ 지점
    Call oForm.DataSources.UserDataSources.Add("DptStr", dt_SHORT_TEXT, 10)     '/ 부서코드
    Call oForm.DataSources.UserDataSources.Add("DptEnd", dt_SHORT_TEXT, 10)
    Call oForm.DataSources.UserDataSources.Add("SMonth", dt_SHORT_TEXT, 2)      '/ 시작월
    Call oForm.DataSources.UserDataSources.Add("EMonth", dt_SHORT_TEXT, 2)      '/ 종료월
    Call oForm.DataSources.UserDataSources.Add("MstCode", dt_SHORT_TEXT, 8)
    Call oForm.DataSources.UserDataSources.Add("MstName", dt_SHORT_TEXT, 30)
    Call oForm.DataSources.UserDataSources.Add("EmpID", dt_SHORT_TEXT, 10)
    Call oForm.DataSources.UserDataSources.Add("Path", dt_SHORT_TEXT)
    Call oForm.DataSources.UserDataSources.Add("STRDAT", dt_DATE)
    Call oForm.DataSources.UserDataSources.Add("ENDDAT", dt_DATE)
    
    Set oEdit = oForm.Items("JsnYear").Specific
    oEdit.DataBind.SetBound True, "", "JsnYear"
    
    Set oEdit = oForm.Items("SMonth").Specific
    oEdit.DataBind.SetBound True, "", "SMonth"
    Set oEdit = oForm.Items("EMonth").Specific
    oEdit.DataBind.SetBound True, "", "EMonth"
    Set oEdit = oForm.Items("MstCode").Specific
    oEdit.DataBind.SetBound True, "", "MstCode"
    Set oEdit = oForm.Items("MstName").Specific
    oEdit.DataBind.SetBound True, "", "MstName"
    Set oEdit = oForm.Items("EmpID").Specific
    oEdit.DataBind.SetBound True, "", "EmpID"
    Set oEdit = oForm.Items("Path").Specific
    oEdit.DataBind.SetBound True, "", "Path"
    Set oEdit = oForm.Items("STRDAT").Specific
    oEdit.DataBind.SetBound True, "", "STRDAT"
    Set oEdit = oForm.Items("ENDDAT").Specific
    oEdit.DataBind.SetBound True, "", "ENDDAT"
    
    '// 생성구분
    Set oCombo1 = oForm.Items("JsnGbn").Specific
    oCombo1.DataBind.SetBound True, "", "JsnGbn"
    oCombo1.ValidValues.Add "1", "연말정산(재직자)"
    oCombo1.ValidValues.Add "2", "중도정산(퇴직자)"
    oCombo1.ValidValues.Add "3", "전체"
    
    oForm.Items("JsnGbn").DisplayDesc = True
    
    '// 사업장
    Set oCombo1 = oForm.Items("CLTCOD").Specific
    oCombo1.DataBind.SetBound True, "", "CLTCOD"
'    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
'    oRecordSet.DoQuery sQry
'    Do Until oRecordSet.EOF
'        oCombo1.ValidValues.Add Trim$(oRecordSet.Fields(0).Value), Trim$(oRecordSet.Fields(1).Value)
'        oRecordSet.MoveNext
'    Loop
    oForm.Items("CLTCOD").DisplayDesc = True
    
'    '// 지점
'    Set oCombo1 = oForm.Items("BPLId").Specific
'    sQry = "SELECT Code, Name FROM OUBR WHERE Code <> '-2' OR (Code = '-2' AND Name <> N'주요') ORDER BY Code ASC"
'    oRecordSet.DoQuery sQry
'    oCombo1.ValidValues.Add "%", "모두"
'    Do Until oRecordSet.EOF
'        oCombo1.ValidValues.Add Trim$(oRecordSet.Fields(0).Value), Trim$(oRecordSet.Fields(1).Value)
'        oRecordSet.MoveNext
'    Loop
'    If oCombo1.ValidValues.Count > 0 Then
'        oCombo1.Select 0, psk_Index
'    End If

    '// 부서
    Set oCombo1 = oForm.Items("DptStr").Specific
    oCombo1.DataBind.SetBound True, "", "DptStr"
    oForm.Items("DptStr").DisplayDesc = True
    '// 부서
    Set oCombo1 = oForm.Items("DptEnd").Specific
    oCombo1.DataBind.SetBound True, "", "DptEnd"
    oForm.Items("DptEnd").DisplayDesc = True
    
    
  '//옵션버튼(생성방법)
    Call oForm.DataSources.UserDataSources.Add("OptionDS", dt_SHORT_TEXT, 1)
    oForm.Items("Opt1").Visible = True
    oForm.Items("Opt2").Visible = True
    Set oOption = oForm.Items("Opt1").Specific
    oOption.DataBind.SetBound True, "", "OptionDS"
    
    Set oOption = oForm.Items("Opt2").Specific
    oOption.GroupWith ("Opt1")
    oOption.DataBind.SetBound True, "", "OptionDS"
    
    oForm.Items("Opt1").Specific.Selected = True
 
    Set oMat1 = oForm.Items("Mat1").Specific
    
    Call oForm.DataSources.UserDataSources.Add("Col1", dt_SHORT_TEXT)
    Set oColumn = oMat1.Columns("Col1")
    oColumn.DataBind.SetBound True, "", "Col1"
    
    oForm.Items("JsnYear").Specific.String = Format$(Now, "YYYY")
    oSMonth = "01"
    oForm.Items("SMonth").Specific.VALUE = oSMonth
    oEMonth = "12"
    oForm.Items("EMonth").Specific.VALUE = oEMonth
    
    Set oEdit = Nothing
    Set oOption = Nothing
    Set oCombo1 = Nothing
    Set oCombo2 = Nothing
    Set oRecordSet = Nothing
    Exit Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Error_Message:
    Set oEdit = Nothing
    Set oOption = Nothing
    Set oCombo1 = Nothing
    Set oCombo2 = Nothing
    Set oRecordSet = Nothing
    Sbo_Application.StatusBar.SetText "CreateItems Error :" & Space$(10) & Err.Description, bmt_Short, smt_Error
End Sub

Private Function Execution_Process()
On Error GoTo Error_Message
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim ErrNum      As Integer
    Dim MSTCNT      As Long
    
    Dim STRDPT      As String
    Dim ENDDPT      As String
    Dim BPLID       As String
    Dim CLTCOD      As String
    Dim STRDAT      As String
    Dim ENDDAT      As String
    ErrNum = 0
    
    '/ 필수Check
    If Trim$(oForm.Items("JsnYear").Specific.String) = "" Then     '/ 정산년도
        ErrNum = 1
        GoTo Error_Message
    ElseIf MDC_SetMod.ChkYearMonth(oJsnYear & Format$(oForm.Items("SMonth").Specific.VALUE, "00")) = False Or _
           MDC_SetMod.ChkYearMonth(oJsnYear & Format$(oForm.Items("EMonth").Specific.VALUE, "00")) = False Then
        ErrNum = 2
        GoTo Error_Message
    ElseIf oForm.Items("SMonth").Specific.VALUE > oForm.Items("EMonth").Specific.VALUE Then
        ErrNum = 3
        GoTo Error_Message
    End If
    oMat1.Clear
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
'/ 임시테이블생성
    sQry = "Exec  ZPY341_1 "
    oRecordSet.DoQuery sQry
   '/ Matrix Message
    oForm.DataSources.UserDataSources("Col1").VALUE = "임시 테이블 생성 완료!"
    oMat1.AddRow
    
    '/ SAP급상여생성 /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    If oForm.DataSources.UserDataSources("OptionDS").ValueEx = "1" Then
'         If oForm.Items("BPLId").Specific.Selected Is Nothing Then
'             ErrNum = 4
'             GoTo Error_Message
        If oForm.Items("DptStr").Specific.Selected Is Nothing Or _
             oForm.Items("DptEnd").Specific.Selected Is Nothing Then
             ErrNum = 5
             GoTo Error_Message
         ElseIf oForm.Items("JsnGbn").Specific.Selected Is Nothing Then
             ErrNum = 7
             GoTo Error_Message
         ElseIf oForm.Items("CLTCOD").Specific.Selected Is Nothing Then
             ErrNum = 8
             GoTo Error_Message
         End If
        oJsnYear = oForm.Items("JsnYear").Specific.String
        oSMonth = oForm.Items("SMonth").Specific.String
        oEMonth = oForm.Items("EMonth").Specific.String
        oJsnGbn = oForm.Items("JsnGbn").Specific.Selected.VALUE
        STRDPT = oForm.Items("DptStr").Specific.Selected.VALUE
        ENDDPT = oForm.Items("DptEnd").Specific.Selected.VALUE
        BPLID = "%"
        CLTCOD = oForm.Items("CLTCOD").Specific.Selected.VALUE
        STRDAT = oForm.Items("STRDAT").Specific.VALUE
        ENDDAT = oForm.Items("ENDDAT").Specific.VALUE
        If Trim$(STRDAT) = "" Or Trim$(ENDDAT) = "" Then
             ErrNum = 9
             GoTo Error_Message
        End If
        If Trim$(ENDDPT) = "%" Then ENDDPT = "ZZZZZZZZ"
        STRDAT = Format$(STRDAT, "0000-00-00")
        ENDDAT = Format$(ENDDAT, "0000-00-00")
    '/ 생성프로시저 실행(년도, 시작월, 종료월, 시작부서, 종료부서, 정산구분, 사원번호
        MSTCNT = 0
        sQry = "Exec ZPY341_3 " & "'" & Trim$(oJsnYear) & "', '" & Trim$(oSMonth) & "','" & Trim$(oEMonth) & "', '" & _
                Trim$(STRDPT) & " ', '" & Trim$(ENDDPT) & "','" & Trim$(oJsnGbn) & "', '" & Trim$(STRDAT) & "',  '" & Trim$(ENDDAT) & "','" & _
                Trim$(CLTCOD) & "', '" & Trim$(BPLID) & "'"
        If Trim$(oForm.Items("MstCode").Specific.String) <> "" Then
            sQry = sQry & ", '" & oForm.Items("MstCode").Specific.String & "'"
        Else
            sQry = sQry & ", '%'"
        End If
        oRecordSet.DoQuery sQry
        '/ Matrix Message
         oForm.DataSources.UserDataSources("Col1").VALUE = "자료 검색중..."
         oMat1.AddRow

        If oRecordSet.RecordCount > 0 Then
            MSTCNT = oRecordSet.Fields(0).VALUE
        End If
        oForm.DataSources.UserDataSources("Col1").VALUE = MSTCNT & "명의 월별 자료 집계 완료."
        oMat1.AddRow
                
        Sbo_Application.StatusBar.SetText "월별 자료 집계 작업을 완료하였습니다.", bmt_Short, smt_Success
    '/ Excel 파일upload /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    ElseIf oForm.DataSources.UserDataSources("OptionDS").ValueEx = "2" Then
         If Trim$(oForm.Items("Path").Specific.String) = "" Then
             ErrNum = 6
             GoTo Error_Message
         End If
         Call Execution_Excel
    End If
 '/
    oForm.DataSources.UserDataSources("Col1").VALUE = "생성된 내용은 [추가]하셔야 시스템에 적용됩니다."
    oMat1.AddRow
            
    
    Set oRecordSet = Nothing
    Execution_Process = True
    Exit Function
'/////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "정산년도를 입력하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 2 Then
        Sbo_Application.StatusBar.SetText "생성기간을 확인하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 3 Then
        Sbo_Application.StatusBar.SetText "시작월보다 종료월이 작습니다. 확인하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 4 Then
        Sbo_Application.StatusBar.SetText "지점을 선택하세요. 선택하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 5 Then
        Sbo_Application.StatusBar.SetText "생성할 부서범위를 선택하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 6 Then
        Sbo_Application.StatusBar.SetText "Upload할 파일을 선택하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 7 Then
        Sbo_Application.StatusBar.SetText "생성구분을 선택하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 8 Then
        Sbo_Application.StatusBar.SetText "자사코드를 선택하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 9 Then
        Sbo_Application.S9tStatusBarMessage "선택 기준일은 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
    Else
        Sbo_Application.StatusBar.SetText "Execution_Process 실행 중 오류가 발생했습니다." & Space$(10) & Err.Description, bmt_Short, smt_Error
    End If
    Execution_Process = False
End Function
Private Function Execution_Save()
On Error GoTo Error_Message
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim ErrNum          As Integer
    Dim JSNYER          As String
    Dim SMonth          As String
    Dim EMonth          As String
'/
    ErrNum = 0
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    sQry = "SELECT TOP 1 * FROM [TmpZPY341]"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
        ErrNum = 1
        GoTo Error_Message
    End If
    JSNYER = oForm.Items("JsnYear").Specific.String
    SMonth = Format$(oForm.Items("SMonth").Specific.VALUE, "00")
    EMonth = Format$(oForm.Items("EMonth").Specific.VALUE, "00")
    '/ 임시테이블저장
    'Exec dbo.ZPY341_4  '2005'
    sQry = "Exec ZPY341_4 " & "'" & Trim$(JSNYER) & "', '" & Trim$(SMonth) & "', '" & Trim$(EMonth) & "', " & oCompany.UserSignature
    oRecordSet.DoQuery sQry

    '/ 메세지
    oForm.DataSources.UserDataSources("Col1").VALUE = "생성된 자료가 시스템에 적용되었습니다!"

    oMat1.AddRow
    
    Set oRecordSet = Nothing

    Execution_Save = True
    Exit Function
'/////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "추가할 자료가 없습니다. 월자료 생성을 먼저 하십시오.", bmt_Short, smt_Error
    Else
        Sbo_Application.StatusBar.SetText "Execution_Save 실행 중 오류가 발생했습니다." & Space$(10) & Err.Description, bmt_Short, smt_Error
    End If
    Execution_Save = False
End Function

Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long)
On Error GoTo Error_Message

    Dim MstInfo         As ZPAY_g_EmpID
        
    Select Case oUID
      Case "JsnYear"
            oJsnYear = oForm.Items(oUID).Specific.String
      Case "SMonth"
            oSMonth = Format$(oForm.Items(oUID).Specific.String, "00")
            oForm.DataSources.UserDataSources("SMonth").ValueEx = oSMonth
      Case "EMonth"
            oEMonth = Format$(oForm.Items(oUID).Specific.String, "00")
            oForm.DataSources.UserDataSources("EMonth").ValueEx = oEMonth
      Case "MstCode"
            If oForm.Items(oUID).Specific.String = "" Then
                oForm.DataSources.UserDataSources(oUID).ValueEx = ""
                oForm.DataSources.UserDataSources("MstName").ValueEx = ""
                oForm.DataSources.UserDataSources("EmpID").ValueEx = ""
            Else
                oForm.DataSources.UserDataSources(oUID).ValueEx = UCase$(oForm.DataSources.UserDataSources(oUID).ValueEx)
                MstInfo = MDC_SetMod.Get_EmpID_InFo(oForm.Items(oUID).Specific.String)
                oForm.DataSources.UserDataSources("MstName").ValueEx = MstInfo.MSTNAM
                oForm.DataSources.UserDataSources("EmpID").ValueEx = MstInfo.EmpID
            End If
            oForm.Items("MstName").Update
            oForm.Items("EmpID").Update
        Case "JsnGbn"
            If Not oForm.Items(oUID).Specific.Selected Is Nothing Then
                oJsnGbn = oForm.Items(oUID).Specific.Selected.VALUE
            Else
                oJsnGbn = ""
            End If
           oForm.Freeze True
           If oJsnYear = "" Then
                oJsnYear = Format$(Now, "YYYY")
                oForm.Items("JsnYear").Specific.VALUE = oJsnYear
           End If
            Select Case Trim$(oJsnGbn)
            Case "2" '/ 중도정산
                oForm.Items("STRDAT").Specific.VALUE = oJsnYear & Format$(Now, "MM" & "01")
                oForm.Items("ENDDAT").Specific.VALUE = oJsnYear & Format$(Now, "MM") & MDC_SetMod.Month_LastDay(oJsnYear & Format$(Now, "MM"))
                oForm.Items("s10").Visible = True
                oForm.Items("s11").Visible = True
                oForm.Items("STRDAT").Visible = True
                oForm.Items("ENDDAT").Visible = True
              
            Case "1" '/ 연말정산
                oForm.Items("STRDAT").Specific.VALUE = oJsnYear & "1231"
                oForm.Items("ENDDAT").Specific.VALUE = oJsnYear & "1231"
                oForm.Items("EMonth").CLICK ct_Regular
                oForm.Items("s10").Visible = True
                oForm.Items("s11").Visible = False
                oForm.Items("STRDAT").Visible = True
                oForm.Items("ENDDAT").Visible = False
            Case "3" '/ 전체
                oForm.Items("STRDAT").Specific.VALUE = oJsnYear & "0101"
                oForm.Items("ENDDAT").Specific.VALUE = oJsnYear & "1231"
                oForm.Items("EMonth").CLICK ct_Regular
                oForm.Items("s10").Visible = False
                oForm.Items("s11").Visible = False
                oForm.Items("STRDAT").Visible = False
                oForm.Items("ENDDAT").Visible = False
            End Select
            oForm.Freeze False
    End Select
    oForm.Items(oUID).Update
    Exit Sub
Error_Message:
    Sbo_Application.StatusBar.SetText "FlushToItemValue Error : " & Err.Description, bmt_Short, smt_Error
End Sub

Private Sub FormItemEnabled()
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim sQry        As String
    Dim i           As Long
    Dim optBtn      As SAPbouiCOM.OptionBtn
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    If oForm.Mode = fm_FIND_MODE Or oForm.Mode = fm_ADD_MODE Then
        oForm.Items("Btn1").Enabled = True
        oForm.Items("Execution").Enabled = True
        
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        
        
        '//부서1
        Set oCombo = oForm.Items("DptStr").Specific
        If oCombo.ValidValues.Count > 0 Then
            For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                oCombo.ValidValues.Remove i, psk_Index
            Next i
            oCombo.ValidValues.Add "", ""
            oCombo.Select 0, psk_Index
        End If
        
        If oForm.DataSources.UserDataSources.Item("CLTCOD").ValueEx <> "" Then
            sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
            sQry = sQry & " WHERE Code = '1' AND U_Char2 = '" & Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) & "'"
            sQry = sQry & " ORDER BY U_Code"
            Call SetReDataCombo(oForm, sQry, oCombo)
            oCombo.ValidValues.Add "%", "전체"
            oCombo.Select "%", psk_ByValue
        End If
        
         '//부서2
        Set oCombo = oForm.Items("DptEnd").Specific
        If oCombo.ValidValues.Count > 0 Then
            For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                oCombo.ValidValues.Remove i, psk_Index
            Next i
            oCombo.ValidValues.Add "", ""
            oCombo.Select 0, psk_Index
        End If
        
        
        If oForm.DataSources.UserDataSources.Item("CLTCOD").ValueEx <> "" Then
            sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
            sQry = sQry & " WHERE Code = '1' AND U_Char2 = '" & Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) & "'"
            sQry = sQry & " ORDER BY U_Code"
            Call SetReDataCombo(oForm, sQry, oCombo)
            oCombo.ValidValues.Add "%", "전체"
            oCombo.Select "%", psk_ByValue
        End If
        
    ElseIf oForm.Mode = fm_OK_MODE Then
         '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD", False)
        
        oForm.Items("Btn1").Enabled = False
        oForm.Items("Execution").Enabled = False
    End If
End Sub
Private Sub Excell_Upload()
On Error GoTo Error_Message
    Dim File_Name As String
    
    File_Name = ZP_Form.OpenDialog(ZP_Form, "*.xls;*.xlsx;*.csv", "파일선택", "C:\")
    If File_Name <> "" Then
        oForm.DataSources.UserDataSources("Path").ValueEx = File_Name
    Else
        oForm.DataSources.UserDataSources("Path").ValueEx = ""
    End If
    oForm.Items("Path").Update
    Exit Sub
'////////////////////////////////////////////////////////////////////
Error_Message:
    Sbo_Application.StatusBar.SetText "Excell_Upload Error:" & Space$(10) & Err.Description, bmt_Short, smt_Error
End Sub

Private Sub Execution_Excel()
On Error GoTo Error_Message
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim ErrNum          As Integer
    
'    Dim x               As Integer
    Dim y               As Integer
    Dim xlApp           As Object
    Dim xlSheet1        As Worksheet
    Dim MSTCOD          As String
    Dim oRow            As Integer
    
    MSTCOD = oForm.Items("MstCode").Specific.String
    ErrNum = 0
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
   
 '/ Open Work Database
    Set xlApp = CreateObject("Excel.Application")
        xlApp.Workbooks.Open oForm.DataSources.UserDataSources("Path").ValueEx
       
    Set xlSheet1 = xlApp.Worksheets(1)
        xlApp.DisplayAlerts = False
    '// 컬럼타이틀값이 현재 패치버전과 맞는지 확인
    If (xlSheet1.Cells(2, 50).Text = "") Or (xlSheet1.Cells(2, 42).Text <> "월지급총액") Then
        ErrNum = 1
        GoTo Error_Message
    ElseIf xlSheet1.Cells(2, 2).Text = "자사코드" And xlSheet1.Cells(3, 2).Text = "" Then
        ErrNum = 2
        GoTo Error_Message
    End If
    y = 3
    oRow = 0
    Do While 1
        tmpArr.JOBDAT = xlSheet1.Cells(y, 1).VALUE          '/ 귀속월
        tmpArr.CLTCOD = xlSheet1.Cells(y, 2).VALUE          '/ 자사코드
        tmpArr.MSTCOD = xlSheet1.Cells(y, 3).VALUE          '/ 사원번호
        tmpArr.MSTNAM = xlSheet1.Cells(y, 4).VALUE          '/ 사원성명
        tmpArr.JIGBIL = xlSheet1.Cells(y, 5).VALUE          '/ 지급일자
        tmpArr.TAMT(1) = Val(xlSheet1.Cells(y, 6).VALUE)    '/ 과세급여
        tmpArr.TAMT(2) = Val(xlSheet1.Cells(y, 7).VALUE)    '/ 과세상여
        tmpArr.TAMT(3) = Val(xlSheet1.Cells(y, 8).VALUE)    '/ 인정상여
        tmpArr.TAMT(4) = Val(xlSheet1.Cells(y, 9).VALUE)    '/ 주식매수행사이익
        tmpArr.TAMT(5) = Val(xlSheet1.Cells(y, 10).VALUE)    '/ 우리사주조합
        tmpArr.TAMT(6) = Val(xlSheet1.Cells(y, 11).VALUE)   '/ 과세총계
        tmpArr.TAMT(7) = Val(xlSheet1.Cells(y, 12).VALUE)   '/ 비과세-G01
        tmpArr.TAMT(8) = Val(xlSheet1.Cells(y, 13).VALUE)   '/ 비과세-H01
        tmpArr.TAMT(9) = Val(xlSheet1.Cells(y, 14).VALUE)   '/ 비과세-H05
        tmpArr.TAMT(10) = Val(xlSheet1.Cells(y, 15).VALUE)  '/ 비과세-H06
        tmpArr.TAMT(11) = Val(xlSheet1.Cells(y, 16).VALUE)  '/ 비과세-H07
        tmpArr.TAMT(12) = Val(xlSheet1.Cells(y, 17).VALUE)  '/ 비과세-H08
        tmpArr.TAMT(13) = Val(xlSheet1.Cells(y, 18).VALUE)  '/ 비과세-H09
        tmpArr.TAMT(14) = Val(xlSheet1.Cells(y, 19).VALUE)  '/ 비과세-H10
        tmpArr.TAMT(15) = Val(xlSheet1.Cells(y, 20).VALUE)  '/ 비과세-H11
        tmpArr.TAMT(16) = Val(xlSheet1.Cells(y, 21).VALUE)  '/ 비과세-H12
        tmpArr.TAMT(17) = Val(xlSheet1.Cells(y, 22).VALUE)  '/ 비과세-H13
        tmpArr.TAMT(18) = Val(xlSheet1.Cells(y, 23).VALUE)  '/ 비과세-I01
        tmpArr.TAMT(19) = Val(xlSheet1.Cells(y, 24).VALUE)  '/ 비과세-K01
        tmpArr.TAMT(20) = Val(xlSheet1.Cells(y, 25).VALUE)  '/ 비과세-M01
        tmpArr.TAMT(21) = Val(xlSheet1.Cells(y, 26).VALUE)  '/ 비과세-M02
        tmpArr.TAMT(22) = Val(xlSheet1.Cells(y, 27).VALUE)  '/ 비과세-M03
        tmpArr.TAMT(23) = Val(xlSheet1.Cells(y, 28).VALUE)  '/ 비과세-O01
        tmpArr.TAMT(24) = Val(xlSheet1.Cells(y, 29).VALUE)  '/ 비과세-Q01
        tmpArr.TAMT(25) = Val(xlSheet1.Cells(y, 30).VALUE)  '/ 비과세-S01
        tmpArr.TAMT(26) = Val(xlSheet1.Cells(y, 31).VALUE)  '/ 비과세-T01
        tmpArr.TAMT(27) = Val(xlSheet1.Cells(y, 32).VALUE)  '/ 비과세-X01
        tmpArr.TAMT(28) = Val(xlSheet1.Cells(y, 33).VALUE)  '/ 비과세-Y01
        tmpArr.TAMT(29) = Val(xlSheet1.Cells(y, 34).VALUE)  '/ 비과세-Y02
        tmpArr.TAMT(30) = Val(xlSheet1.Cells(y, 35).VALUE)  '/ 비과세-Y03
        tmpArr.TAMT(31) = Val(xlSheet1.Cells(y, 36).VALUE)  '/ 비과세-Y20
        tmpArr.TAMT(32) = Val(xlSheet1.Cells(y, 37).VALUE)  '/ 비과세-Y20
        tmpArr.TAMT(33) = Val(xlSheet1.Cells(y, 38).VALUE)  '/ 비과세-Z01
        tmpArr.TAMT(34) = Val(xlSheet1.Cells(y, 39).VALUE)  '/ 비과세-차량,식대보조        BIGWA02
        tmpArr.TAMT(35) = Val(xlSheet1.Cells(y, 40).VALUE)  '/ 비과세-기타(지급조서포함)   BIGWA04
        tmpArr.TAMT(36) = Val(xlSheet1.Cells(y, 41).VALUE)  '/ 비과세-기타(지급조서미포함) BIGWA07
        tmpArr.TAMT(37) = Val(xlSheet1.Cells(y, 42).VALUE)  '/ 월지급총액
        tmpArr.TAMT(38) = Val(xlSheet1.Cells(y, 43).VALUE)  '/ 국민연금
        tmpArr.TAMT(39) = Val(xlSheet1.Cells(y, 44).VALUE)  '/ 건강보험
        tmpArr.TAMT(40) = Val(xlSheet1.Cells(y, 45).VALUE)  '/ 고용보험
        tmpArr.TAMT(41) = Val(xlSheet1.Cells(y, 46).VALUE)  '/ 소득세
        tmpArr.TAMT(42) = Val(xlSheet1.Cells(y, 47).VALUE)  '/ 주민세
        tmpArr.TAMT(43) = Val(xlSheet1.Cells(y, 48).VALUE)  '/ 농특세
        tmpArr.TAMT(44) = Val(xlSheet1.Cells(y, 49).VALUE)  '/ 기부금
        tmpArr.TAMT(45) = Val(xlSheet1.Cells(y, 50).VALUE)  '/ 장기요양보험
        
        If Trim$(tmpArr.JOBDAT) = "" Or Trim$(tmpArr.MSTCOD) = "" Or Trim$(tmpArr.CLTCOD) = "" Then
            Exit Do
            ErrNum = 3
            GoTo Error_Message
        End If
        If Trim$(MSTCOD) = "" Or (Trim$(MSTCOD) <> "" And Trim$(tmpArr.MSTCOD) = Trim$(MSTCOD)) Then
            '/ 프로시저실행
            sQry = "Exec ZPY341_2 '" & tmpArr.JOBDAT & "','" & Trim$(tmpArr.CLTCOD) & "','" & Trim$(tmpArr.MSTCOD) & "','" & tmpArr.JIGBIL & "'," & _
                    tmpArr.TAMT(1) & "," & tmpArr.TAMT(2) & "," & tmpArr.TAMT(3) & "," & tmpArr.TAMT(4) & "," & _
                    tmpArr.TAMT(5) & "," & tmpArr.TAMT(6) & "," & tmpArr.TAMT(7) & "," & tmpArr.TAMT(8) & "," & _
                    tmpArr.TAMT(9) & "," & tmpArr.TAMT(10) & "," & tmpArr.TAMT(11) & "," & tmpArr.TAMT(12) & "," & _
                    tmpArr.TAMT(13) & "," & tmpArr.TAMT(14) & "," & tmpArr.TAMT(15) & "," & tmpArr.TAMT(16) & "," & _
                    tmpArr.TAMT(17) & "," & tmpArr.TAMT(18) & "," & tmpArr.TAMT(19) & "," & tmpArr.TAMT(20) & "," & _
                    tmpArr.TAMT(21) & "," & tmpArr.TAMT(22) & "," & tmpArr.TAMT(23) & "," & tmpArr.TAMT(24) & "," & _
                    tmpArr.TAMT(25) & "," & tmpArr.TAMT(26) & "," & tmpArr.TAMT(27) & "," & tmpArr.TAMT(28) & "," & _
                    tmpArr.TAMT(29) & "," & tmpArr.TAMT(30) & "," & tmpArr.TAMT(31) & "," & tmpArr.TAMT(32) & "," & _
                    tmpArr.TAMT(33) & "," & tmpArr.TAMT(34) & "," & tmpArr.TAMT(35) & "," & tmpArr.TAMT(36) & "," & _
                    tmpArr.TAMT(37) & "," & tmpArr.TAMT(38) & "," & tmpArr.TAMT(39) & "," & tmpArr.TAMT(40) & "," & _
                    tmpArr.TAMT(41) & "," & tmpArr.TAMT(42) & "," & tmpArr.TAMT(43) & "," & tmpArr.TAMT(44) & "," & _
                    tmpArr.TAMT(45)
                    
            oRecordSet.DoQuery sQry
            If oRecordSet.RecordCount = 0 Then
                oForm.DataSources.UserDataSources("Col1").VALUE = "행: (" & y & ") 사원번호:" & tmpArr.MSTCOD & "일치하는 사원이 없습니다."
                oMat1.AddRow
            End If
            oRow = oRow + 1
        End If
        y = y + 1
    Loop
    oForm.Update
    oForm.DataSources.UserDataSources("Col1").VALUE = "엑셀 upload Line(" & oRow & "/ " & y - 3 & ")작업을 완료하였습니다."
    oMat1.AddRow
    
    Sbo_Application.StatusBar.SetText "작업을 완료하였습니다.", bmt_Short, smt_Success
       xlApp.Quit
   Set xlSheet1 = Nothing
   Set xlApp = Nothing
   Exit Sub
'/ Message /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
Error_Message:
On Error Resume Next
   Screen.MousePointer = vbNormal
   
       xlApp.DisplayAlerts = False
       xlApp.Quit
   Set xlSheet1 = Nothing
   Set xlApp = Nothing
    If ErrNum = "1" Then
        Sbo_Application.StatusBar.SetText "엑셀서식이 패치버전과 일치하지 않습니다. 관리자에게 문의하십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = "2" Then
        Sbo_Application.StatusBar.SetText "자사코드는 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = "3" Then
        Sbo_Application.StatusBar.SetText y & "번째 행의 자사코드,사원번호, 귀속연월는 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
    Else
        Sbo_Application.StatusBar.SetText "엑셀 작업에 실패 하였습니다." & Space$(10) & Err.Description, bmt_Short, smt_Error
    End If
End Sub
Private Sub TmpTable_Delete()
On Error GoTo Error_Message
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
'/ 임시테이블생성
    sQry = "Exec  ZPY341_1 "
    oRecordSet.DoQuery sQry
    
    Set oRecordSet = Nothing
    Exit Sub
'/////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set oRecordSet = Nothing
    Sbo_Application.StatusBar.SetText "TmpTable_Delete 실행중 오류. 프로시저가 있는지 확인하세요." & Space$(10) & Err.Description, bmt_Short, smt_Error
End Sub

