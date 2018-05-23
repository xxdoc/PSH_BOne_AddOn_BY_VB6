VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY401.cls
'//  Module         : 인사관리 > 연말정산관리
'//  Desc           : 전근무지 등록
'//********************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form

Public oGrid1           As SAPbouiCOM.Grid
Public oDS_PH_PY401A     As SAPbouiCOM.DataTable


Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long

Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY401.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY401_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY401"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
'    oForm.DataBrowser.BrowseBy = "Code"
        
    oForm.PaneLevel = 1
    oForm.Freeze True
    Call PH_PY401_CreateItems
    Call PH_PY401_FormItemEnabled
    Call PH_PY401_EnableMenus
'    Call PH_PY401_SetDocument(oFromDocEntry01)
'    Call PH_PY401_FormResize
    
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

Private Function PH_PY401_CreateItems() As Boolean

    Dim sQry        As String
    Dim i           As Long
    Dim CLTCOD      As String
    
    Dim oCheck      As SAPbouiCOM.CheckBox
    Dim oEdit       As SAPbouiCOM.EditText
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns    As SAPbouiCOM.Columns
    Dim optBtn      As SAPbouiCOM.OptionBtn
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY401_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Set oGrid1 = oForm.Items("Grid01").Specific
    
    oForm.DataSources.DataTables.Add ("PH_PY401")
    
    oGrid1.DataTable = oForm.DataSources.DataTables.Item("PH_PY401")
    Set oDS_PH_PY401A = oForm.DataSources.DataTables.Item("PH_PY401")
    

    '//----------------------------------------------------------------------------------------------
    '// 기본사항
    '//----------------------------------------------------------------------------------------------
    
    '//사업장
    
    Set oCombo = oForm.Items("CLTCOD").Specific
'    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
'    Call SetReDataCombo(oForm, sQry, oCombo)
'
'    CLTCOD = MDC_SetMod.Get_ReData("Branch", "USER_CODE", "OUSR", "'" & oCompany.UserName & "'")
'    oCombo.Select CLTCOD, psk_ByValue
    
       
'    oCombo.Select 0, psk_Index
    oForm.Items("CLTCOD").DisplayDesc = True
    
    '//년도
    Call oForm.DataSources.UserDataSources.Add("Year", dt_SHORT_TEXT, 4)
    oForm.Items("Year").Specific.DataBind.SetBound True, "", "Year"
    
    '//사번
    Call oForm.DataSources.UserDataSources.Add("MSTCOD", dt_SHORT_TEXT, 10)
    oForm.Items("MSTCOD").Specific.DataBind.SetBound True, "", "MSTCOD"
    '//성명
    Call oForm.DataSources.UserDataSources.Add("FullName", dt_SHORT_TEXT, 20)
    oForm.Items("FullName").Specific.DataBind.SetBound True, "", "FullName"
    
    '//사업자등록번호1
    Call oForm.DataSources.UserDataSources.Add("entno1", dt_SHORT_TEXT, 20)
    oForm.Items("entno1").Specific.DataBind.SetBound True, "", "entno1"
    '//사업자등록번호2
    Call oForm.DataSources.UserDataSources.Add("entno2", dt_SHORT_TEXT, 20)
    oForm.Items("entno2").Specific.DataBind.SetBound True, "", "entno2"
    
    '//근무처명1
    Call oForm.DataSources.UserDataSources.Add("servcomp1", dt_SHORT_TEXT, 30)
    oForm.Items("servcomp1").Specific.DataBind.SetBound True, "", "servcomp1"
    '//근무처명2
    Call oForm.DataSources.UserDataSources.Add("servcomp2", dt_SHORT_TEXT, 30)
    oForm.Items("servcomp2").Specific.DataBind.SetBound True, "", "servcomp2"
    
    '//시작일1
    Call oForm.DataSources.UserDataSources.Add("symd1", dt_DATE, 8)
    oForm.Items("symd1").Specific.DataBind.SetBound True, "", "symd1"
    '//시작일2
    Call oForm.DataSources.UserDataSources.Add("symd2", dt_DATE, 8)
    oForm.Items("symd2").Specific.DataBind.SetBound True, "", "symd2"
    
    '//종료일1
    Call oForm.DataSources.UserDataSources.Add("eymd1", dt_DATE, 8)
    oForm.Items("eymd1").Specific.DataBind.SetBound True, "", "eymd1"
    '//종료일2
    Call oForm.DataSources.UserDataSources.Add("eymd2", dt_DATE, 8)
    oForm.Items("eymd2").Specific.DataBind.SetBound True, "", "eymd2"
    
    '//급여총액1
    Call oForm.DataSources.UserDataSources.Add("payrtot1", dt_SUM)
    oForm.Items("payrtot1").Specific.DataBind.SetBound True, "", "payrtot1"
    '//급여총액2
    Call oForm.DataSources.UserDataSources.Add("payrtot2", dt_SUM)
    oForm.Items("payrtot2").Specific.DataBind.SetBound True, "", "payrtot2"
    
    '//상여총액1
    Call oForm.DataSources.UserDataSources.Add("bnstot1", dt_SUM)
    oForm.Items("bnstot1").Specific.DataBind.SetBound True, "", "bnstot1"
    '//상여총액2
    Call oForm.DataSources.UserDataSources.Add("bnstot2", dt_SUM)
    oForm.Items("bnstot2").Specific.DataBind.SetBound True, "", "bnstot2"
    
    '//국외근로1
    Call oForm.DataSources.UserDataSources.Add("fwork1", dt_SUM)
    oForm.Items("fwork1").Specific.DataBind.SetBound True, "", "fwork1"
    '//국외근로2
    Call oForm.DataSources.UserDataSources.Add("fwork2", dt_SUM)
    oForm.Items("fwork2").Specific.DataBind.SetBound True, "", "fwork2"
    
    '//야간근로수당등1
    Call oForm.DataSources.UserDataSources.Add("ndtalw1", dt_SUM)
    oForm.Items("ndtalw1").Specific.DataBind.SetBound True, "", "ndtalw1"
    '//야간근로수당등2
    Call oForm.DataSources.UserDataSources.Add("ndtalw2", dt_SUM)
    oForm.Items("ndtalw2").Specific.DataBind.SetBound True, "", "ndtalw2"
    
    '//기타비과세1
    Call oForm.DataSources.UserDataSources.Add("etcntax1", dt_SUM)
    oForm.Items("etcntax1").Specific.DataBind.SetBound True, "", "etcntax1"
    '//기타비과세2
    Call oForm.DataSources.UserDataSources.Add("etcntax2", dt_SUM)
    oForm.Items("etcntax2").Specific.DataBind.SetBound True, "", "etcntax2"
    
    '//중식수당1
    Call oForm.DataSources.UserDataSources.Add("lnchalw1", dt_SUM)
    oForm.Items("lnchalw1").Specific.DataBind.SetBound True, "", "lnchalw1"
    '//중식수당2
    Call oForm.DataSources.UserDataSources.Add("lnchalw2", dt_SUM)
    oForm.Items("lnchalw2").Specific.DataBind.SetBound True, "", "lnchalw2"
    
    '//외국납부세액1
    Call oForm.DataSources.UserDataSources.Add("ftaxamt1", dt_SUM)
    oForm.Items("ftaxamt1").Specific.DataBind.SetBound True, "", "ftaxamt1"
    '//외국납부세액2
    Call oForm.DataSources.UserDataSources.Add("ftaxamt2", dt_SUM)
    oForm.Items("ftaxamt2").Specific.DataBind.SetBound True, "", "ftaxamt2"
    
    '//저축세액공제1
    Call oForm.DataSources.UserDataSources.Add("savtaxddc1", dt_SUM)
    oForm.Items("savtaxddc1").Specific.DataBind.SetBound True, "", "savtaxddc1"
    '//저축세액공제2
    Call oForm.DataSources.UserDataSources.Add("savtaxddc2", dt_SUM)
    oForm.Items("savtaxddc2").Specific.DataBind.SetBound True, "", "savtaxddc2"
    
    '//소득세1
    Call oForm.DataSources.UserDataSources.Add("incmtax1", dt_SUM)
    oForm.Items("incmtax1").Specific.DataBind.SetBound True, "", "incmtax1"
    '//소득세2
    Call oForm.DataSources.UserDataSources.Add("indmtax2", dt_SUM)
    oForm.Items("indmtax2").Specific.DataBind.SetBound True, "", "indmtax2"
    
    '//농어촌특별세1
    Call oForm.DataSources.UserDataSources.Add("fvsptax1", dt_SUM)
    oForm.Items("fvsptax1").Specific.DataBind.SetBound True, "", "fvsptax1"
    '//농어촌특별세2
    Call oForm.DataSources.UserDataSources.Add("fvsptax2", dt_SUM)
    oForm.Items("fvsptax2").Specific.DataBind.SetBound True, "", "fvsptax2"
    
    '//주민세1
    Call oForm.DataSources.UserDataSources.Add("residtax1", dt_SUM)
    oForm.Items("residtax1").Specific.DataBind.SetBound True, "", "residtax1"
    '//주민세2
    Call oForm.DataSources.UserDataSources.Add("residtax2", dt_SUM)
    oForm.Items("residtax2").Specific.DataBind.SetBound True, "", "residtax2"
    
    '//건강보험1
    Call oForm.DataSources.UserDataSources.Add("medcinsr1", dt_SUM)
    oForm.Items("medcinsr1").Specific.DataBind.SetBound True, "", "medcinsr1"
    '//건강보험2
    Call oForm.DataSources.UserDataSources.Add("medcinsr2", dt_SUM)
    oForm.Items("medcinsr2").Specific.DataBind.SetBound True, "", "medcinsr2"
    
    '//고용보험1
    Call oForm.DataSources.UserDataSources.Add("asopinsr1", dt_SUM)
    oForm.Items("asopinsr1").Specific.DataBind.SetBound True, "", "asopinsr1"
    '//고용보험2
    Call oForm.DataSources.UserDataSources.Add("asopinsr2", dt_SUM)
    oForm.Items("asopinsr2").Specific.DataBind.SetBound True, "", "asopinsr2"
    
    '//연금보험료공제1
    Call oForm.DataSources.UserDataSources.Add("annuboamt1", dt_SUM)
    oForm.Items("annuboamt1").Specific.DataBind.SetBound True, "", "annuboamt1"
    '//고용보험2
    Call oForm.DataSources.UserDataSources.Add("annuboamt2", dt_SUM)
    oForm.Items("annuboamt2").Specific.DataBind.SetBound True, "", "annuboamt2"
    
    
    oForm.Update
    
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
    
PH_PY401_CreateItems_Error:
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY401_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY401_EnableMenus()
    
    On Error GoTo PH_PY401_EnableMenus_Error

    Call oForm.EnableMenu("1283", False)     '//제거
    Call oForm.EnableMenu("1284", False)    '//취소
    Call oForm.EnableMenu("1293", False)     '//행삭제
    
    Exit Sub
    
PH_PY401_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY401_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY401_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY401_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY401_FormItemEnabled
'        Call PH_PY401_AddMatrixRow
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY401_FormItemEnabled
        oForm.Items("Code").Specific.VALUE = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY401_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY401_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY401_FormItemEnabled()
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim sQry        As String
    Dim i           As Long
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    
    Dim CLTCOD As String
    Dim sPosDate As String
    
    On Error GoTo PH_PY401_FormItemEnabled_Error
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Call oForm.Freeze(True)
    If (oForm.Mode = fm_ADD_MODE) Then
    
        
        Call oForm.EnableMenu("1281", False)     '//문서찾기
        Call oForm.EnableMenu("1282", True)    '//문서추가
        
        oForm.Items("Year").Specific.VALUE = Format(Now, "YYYY") - 1
        oForm.Items("MSTCOD").Specific.VALUE = ""
        oForm.Items("FullName").Specific.VALUE = ""
        
        oForm.DataSources.UserDataSources.Item("entno1").VALUE = ""
        oForm.DataSources.UserDataSources.Item("servcomp1").VALUE = ""
        oForm.DataSources.UserDataSources.Item("symd1").VALUE = ""
        oForm.DataSources.UserDataSources.Item("eymd1").VALUE = ""

        oForm.DataSources.UserDataSources.Item("payrtot1").VALUE = 0
        oForm.DataSources.UserDataSources.Item("bnstot1").VALUE = 0
        oForm.DataSources.UserDataSources.Item("fwork1").VALUE = 0
        oForm.DataSources.UserDataSources.Item("ndtalw1").VALUE = 0
        oForm.DataSources.UserDataSources.Item("etcntax1").VALUE = 0
        oForm.DataSources.UserDataSources.Item("lnchalw1").VALUE = 0
        oForm.DataSources.UserDataSources.Item("ftaxamt1").VALUE = 0
        oForm.DataSources.UserDataSources.Item("savtaxddc1").VALUE = 0
        oForm.DataSources.UserDataSources.Item("incmtax1").VALUE = 0
        oForm.DataSources.UserDataSources.Item("fvsptax1").VALUE = 0
        oForm.DataSources.UserDataSources.Item("residtax1").VALUE = 0
        oForm.DataSources.UserDataSources.Item("medcinsr1").VALUE = 0
        oForm.DataSources.UserDataSources.Item("asopinsr1").VALUE = 0
        oForm.DataSources.UserDataSources.Item("annuboamt1").VALUE = 0


        oForm.DataSources.UserDataSources.Item("entno2").VALUE = ""
        oForm.DataSources.UserDataSources.Item("servcomp2").VALUE = ""
        oForm.DataSources.UserDataSources.Item("symd2").VALUE = ""
        oForm.DataSources.UserDataSources.Item("eymd2").VALUE = ""

        oForm.DataSources.UserDataSources.Item("payrtot2").VALUE = 0
        oForm.DataSources.UserDataSources.Item("bnstot2").VALUE = 0
        oForm.DataSources.UserDataSources.Item("fwork2").VALUE = 0
        oForm.DataSources.UserDataSources.Item("ndtalw2").VALUE = 0
        oForm.DataSources.UserDataSources.Item("etcntax2").VALUE = 0
        oForm.DataSources.UserDataSources.Item("lnchalw2").VALUE = 0
        oForm.DataSources.UserDataSources.Item("ftaxamt2").VALUE = 0
        oForm.DataSources.UserDataSources.Item("savtaxddc2").VALUE = 0
        oForm.DataSources.UserDataSources.Item("indmtax2").VALUE = 0
        oForm.DataSources.UserDataSources.Item("fvsptax2").VALUE = 0
        oForm.DataSources.UserDataSources.Item("residtax2").VALUE = 0
        oForm.DataSources.UserDataSources.Item("medcinsr2").VALUE = 0
        oForm.DataSources.UserDataSources.Item("asopinsr2").VALUE = 0
        oForm.DataSources.UserDataSources.Item("annuboamt2").VALUE = 0
        
        oForm.Items("TeamName").Specific.VALUE = ""
        oForm.Items("RspName").Specific.VALUE = ""
        oForm.Items("ClsName").Specific.VALUE = ""
        
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")


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
    Set oCombo = Nothing
    Call oForm.Freeze(False)
    Exit Sub
    
PH_PY401_FormItemEnabled_Error:
    Set oCombo = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY401_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Public Sub Raise_FormItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
    Dim sQry        As String
    Dim i           As Long
    Dim tSex        As String
    Dim tBrith      As String
    Dim Day         As String
    Dim ActCode     As String
    Dim CLTCOD      As String
    Dim MSTCOD      As String
    Dim yyyy        As String
    Dim Result      As String
    
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns     As SAPbouiCOM.Columns
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo Raise_FormItemEvent_Error
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Select Case pval.EventType
        Case et_ITEM_PRESSED: '//1
            
            If pval.BeforeAction = True Then
                If pval.ItemUID = "1" Then
                    If PH_PY401_DataValidCheck = False Then
                        BubbleEvent = False
                    End If
                End If
                
                If pval.ItemUID = "Btn_ret" Then
                    Call PH_PY401_MTX01
                End If
                
                
                
                If pval.ItemUID = "Btn01" Then
                
'                    CLTCOD = oForm.Items("CLTCOD").Specific.VALUE
'                    yyyy = oForm.Items("Year").Specific.VALUE
'                    sQry = " Select CODE from [@PSH_PAYBACK] WHERE CODE = '" & yyyy & CLTCOD & "'"
'                    oRecordSet.DoQuery sQry
'
'                    Result = oRecordSet.Fields(0).VALUE
'                    If Result <> "" Then
'                        Call MDC_Com.MDC_GF_Message("등록불가 년도입니다. 담당자에게 문의바랍니다.", "E")
'                    End If
'                    If Result = "" Then
                        Call PH_PY401_SAVE
'                    End If
                End If
                
                                
                If pval.ItemUID = "Btn_del" Then
                    Call PH_PY401_Delete
                    Call PH_PY401_FormItemEnabled
                End If
'                If oForm.Mode = fm_FIND_MODE Then
'                    If pval.ItemUID = "Btn01" Then
'                        Sbo_Application.ActivateMenuItem ("7425")
'                        BubbleEvent = False
'                    End If
'
'                End If
            ElseIf pval.BeforeAction = False Then
                Select Case pval.ItemUID
                Case "1"
                    If oForm.Mode = fm_ADD_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY401_FormItemEnabled
                        End If
                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY401_FormItemEnabled
                        End If
                    ElseIf oForm.Mode = fm_OK_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY401_FormItemEnabled
                        End If
                    End If
'
                End Select
            End If
     
'----------------------------------------------------------
        Case et_KEY_DOWN: '//2
            If pval.BeforeAction = True Then
                If pval.CharPressed = 9 Then
                    If pval.ItemUID = "MSTCOD" Then
                        If oForm.Items("MSTCOD").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                                BubbleEvent = False
                        End If
                    End If
                End If
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
                    '//사업장(헤더)
                    If pval.ItemUID = "SCLTCOD" Then
                        
                    End If
                    
                End If
            End If

            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_CLICK: '//6
            Call oForm.Freeze(True)
            If pval.BeforeAction = True Then
                Select Case pval.ItemUID
                Case "Grid01"
                    If pval.Row >= 0 Then
                        Select Case pval.ItemUID
                        Case "Grid01"
                            'Call oMat1.SelectRow(pval.Row, True, False)
                            PH_PY401_MTX02 pval.ItemUID, pval.Row, pval.ColUID
                        End Select
                        
                    End If
                End Select
                
                Select Case pval.ItemUID
                Case "Grid01"
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
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_DOUBLE_CLICK: '//7
            Call oForm.Freeze(True)
            If pval.BeforeAction = True Then
            Else
                
            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_MATRIX_LINK_PRESSED '//8

'----------------------------------------------------------
        Case et_MATRIX_COLLAPSE_PRESSED '//9
'----------------------------------------------------------
        Case et_VALIDATE: '//10
'            Call oForm.Freeze(True)
            If pval.BeforeAction = True Then
                If pval.ItemChanged = True Then
                    
                End If
                    
            ElseIf pval.BeforeAction = False Then
                If pval.ItemChanged = True Then
                    Select Case pval.ItemUID

                        Case "MSTCOD"
                            CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE)
                            MSTCOD = oForm.Items("MSTCOD").Specific.VALUE
                            
                            sQry = "Select Code,"
                            sQry = sQry + " FullName = U_FullName,"
                            sQry = sQry + " TeamName = Isnull((SELECT U_CodeNm"
                            sQry = sQry + " From [@PS_HR200L]"
                            sQry = sQry + " WHERE Code = '1'"
                            sQry = sQry + " And U_Code = U_TeamCode),''),"
                            sQry = sQry + " RspName  = Isnull((SELECT U_CodeNm"
                            sQry = sQry + " From [@PS_HR200L]"
                            sQry = sQry + " WHERE Code = '2'"
                            sQry = sQry + " And U_Code = U_RspCode),''),"
                            sQry = sQry + " ClsName  = Isnull((SELECT U_CodeNm"
                            sQry = sQry + " From [@PS_HR200L]"
                            sQry = sQry + " WHERE Code = '9'"
                            sQry = sQry + " And U_Code  = U_ClsCode"
                            sQry = sQry + " And U_Char3 = U_CLTCOD),'')"
                            sQry = sQry + " From [@PH_PY001A]"
                            sQry = sQry + " Where U_CLTCOD = '" + CLTCOD + "'"
                            sQry = sQry + " and Code = '" + MSTCOD + "'"
                              
                            oRecordSet.DoQuery sQry
                            
                            oForm.Items("FullName").Specific.VALUE = oRecordSet.Fields("FullName").VALUE
                            oForm.Items("TeamName").Specific.VALUE = oRecordSet.Fields("TeamName").VALUE
                            oForm.Items("RspName").Specific.VALUE = oRecordSet.Fields("RspName").VALUE
                            oForm.Items("ClsName").Specific.VALUE = oRecordSet.Fields("ClsName").VALUE
                        
                    End Select
                 
                End If
            End If
'            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_MATRIX_LOAD: '//11
            If pval.BeforeAction = True Then
            ElseIf pval.BeforeAction = False Then
'                oMat1.LoadFromDataSource
'                Call PH_PY401_AddMatrixRow
                
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
                Set oDS_PH_PY401A = Nothing
                
'                Set oMat1 = Nothing
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
'                If pval.ItemUID = "Code" Then
'                    Call MDC_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY401A", "Code")
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
    Sbo_Application.SetStatusBarMessage "Raise_FormItemEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
'                Call PH_PY401_FormItemEnabled
        End Select
    ElseIf (pval.BeforeAction = False) Then
        Select Case pval.MenuUID
            Case "1283":
                oForm.Mode = fm_ADD_MODE
                Call PH_PY401_FormItemEnabled
'                Call PH_PY401_AddMatrixRow
            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//문서찾기
                Call PH_PY401_FormItemEnabled
'                Call PH_PY401_AddMatrixRow
                oForm.Items("Code").CLICK ct_Regular
            Case "1282":    '//문서추가
                Call PH_PY401_FormItemEnabled
'                Call PH_PY401_AddMatrixRow
            Case "1288", "1289", "1290", "1291":
                Call PH_PY401_FormItemEnabled
            Case "1293" '// 행삭제
                
        End Select
    End If
    Call oForm.Freeze(False)
    Exit Sub
Raise_FormMenuEvent_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_MenuEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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


Sub PH_PY401_FormClear()
On Error GoTo PH_PY401_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY401'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PH_PY401_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY401_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY401_DataValidCheck() As Boolean
On Error GoTo PH_PY401_DataValidCheck_Error
    PH_PY401_DataValidCheck = False
    Dim i       As Long
    Dim j       As Long
    
    Dim sQry    As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Exit Function
    
    Set oRecordSet = Nothing
    
    
PH_PY401_DataValidCheck_Error:
    Set oRecordSet = Nothing
    PH_PY401_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY401_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY401_MTX01()

    '//메트릭스에 데이터 로드
    
    Dim i       As Long
    Dim sQry    As String
    Dim iRow        As Long
    
    Dim Param01 As String
    Dim Param02 As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY401_MTX01_Error

    Call oForm.Freeze(True)
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    Param01 = Trim(oForm.Items("CLTCOD").Specific.VALUE)
    Param02 = oForm.Items("Year").Specific.VALUE
        
    If Trim(Param01) = "" Then
        Call MDC_Com.MDC_GF_Message("사업장이 없습니다. 확인바랍니다..", "E")
        GoTo PH_PY401_MTX01_Exit
    End If
        
    If Trim(Param02) = "" Then
        Call MDC_Com.MDC_GF_Message("년도가 없습니다. 확인바랍니다..", "E")
        GoTo PH_PY401_MTX01_Exit
    End If
        
        
        
    sQry = "EXEC PH_PY401_01 '" & Param01 & "', '" & Param02 & "'"

    Call oDS_PH_PY401A.ExecuteQuery(sQry)
    
    
    
    iRow = oForm.DataSources.DataTables.Item(0).Rows.Count
    
    Call PH_PY401_TitleSetting(iRow)
    
    oForm.Update
    
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Sub
PH_PY401_MTX01_Exit:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Sub
PH_PY401_MTX01_Error:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY401_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub
Private Sub PH_PY401_MTX02(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)


    '//그리드 자료를 head에 로드
    
    Dim i       As Long
    Dim sQry    As String
    Dim sRow As Long
    
    Dim Param01 As String
    Dim Param02 As String
    Dim Param03 As String
    
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY401_MTX02_Error

    Call oForm.Freeze(True)
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    sRow = oRow
    
    
    Param01 = Trim(oForm.Items("CLTCOD").Specific.VALUE)
    Param02 = oDS_PH_PY401A.Columns.Item("Year").Cells(oRow).VALUE
    Param03 = oDS_PH_PY401A.Columns.Item("MSTCOD").Cells(oRow).VALUE
    
    
    
    sQry = "EXEC PH_PY401_02 '" & Param01 & "', '" & Param02 & "', '" & Param03 & "'"
    Call oRecordSet.DoQuery(sQry)
    
    If (oRecordSet.RecordCount = 0) Then
        
        oForm.Items("MSTCOD").Specific.VALUE = oDS_PH_PY401A.Columns.Item("MSTCOD").Cells(oRow).VALUE
        oForm.Items("FullName").Specific.VALUE = oDS_PH_PY401A.Columns.Item("FullName").Cells(oRow).VALUE
        
        oForm.DataSources.UserDataSources.Item("entno1").VALUE = ""
        oForm.DataSources.UserDataSources.Item("servcomp1").VALUE = ""
        oForm.DataSources.UserDataSources.Item("symd1").VALUE = ""
        oForm.DataSources.UserDataSources.Item("eymd1").VALUE = ""
        
        oForm.DataSources.UserDataSources.Item("payrtot1").VALUE = 0
        oForm.DataSources.UserDataSources.Item("bnstot1").VALUE = 0
        oForm.DataSources.UserDataSources.Item("fwork1").VALUE = 0
        oForm.DataSources.UserDataSources.Item("ndtalw1").VALUE = 0
        oForm.DataSources.UserDataSources.Item("etcntax1").VALUE = 0
        oForm.DataSources.UserDataSources.Item("lnchalw1").VALUE = 0
        oForm.DataSources.UserDataSources.Item("ftaxamt1").VALUE = 0
        oForm.DataSources.UserDataSources.Item("savtaxddc1").VALUE = 0
        oForm.DataSources.UserDataSources.Item("incmtax1").VALUE = 0
        oForm.DataSources.UserDataSources.Item("fvsptax1").VALUE = 0
        oForm.DataSources.UserDataSources.Item("residtax1").VALUE = 0
        oForm.DataSources.UserDataSources.Item("medcinsr1").VALUE = 0
        oForm.DataSources.UserDataSources.Item("asopinsr1").VALUE = 0
        oForm.DataSources.UserDataSources.Item("annuboamt1").VALUE = 0
        
        
        oForm.DataSources.UserDataSources.Item("entno2").VALUE = ""
        oForm.DataSources.UserDataSources.Item("servcomp2").VALUE = ""
        oForm.DataSources.UserDataSources.Item("symd2").VALUE = ""
        oForm.DataSources.UserDataSources.Item("eymd2").VALUE = ""
        
        oForm.DataSources.UserDataSources.Item("payrtot2").VALUE = 0
        oForm.DataSources.UserDataSources.Item("bnstot2").VALUE = 0
        oForm.DataSources.UserDataSources.Item("fwork2").VALUE = 0
        oForm.DataSources.UserDataSources.Item("ndtalw2").VALUE = 0
        oForm.DataSources.UserDataSources.Item("etcntax2").VALUE = 0
        oForm.DataSources.UserDataSources.Item("lnchalw2").VALUE = 0
        oForm.DataSources.UserDataSources.Item("ftaxamt2").VALUE = 0
        oForm.DataSources.UserDataSources.Item("savtaxddc2").VALUE = 0
        oForm.DataSources.UserDataSources.Item("indmtax2").VALUE = 0
        oForm.DataSources.UserDataSources.Item("fvsptax2").VALUE = 0
        oForm.DataSources.UserDataSources.Item("residtax2").VALUE = 0
        oForm.DataSources.UserDataSources.Item("medcinsr2").VALUE = 0
        oForm.DataSources.UserDataSources.Item("asopinsr2").VALUE = 0
        oForm.DataSources.UserDataSources.Item("annuboamt2").VALUE = 0
        
        
        
'        oForm.Items("TeamName").Specific.VALUE = oRecordSet.Fields("TeamName").VALUE
'        oForm.Items("RspName").Specific.VALUE = oRecordSet.Fields("RspName").VALUE
'        oForm.Items("ClsName").Specific.VALUE = oRecordSet.Fields("ClsName").VALUE
    
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
        GoTo PH_PY401_MTX02_Exit
    End If
    
    
'    oForm.Items("PosDate").Specific.Value = Format(oRecordset.Fields("PosDate").Value, "YYYYMMDD")
        
'    oForm.DataSources.UserDataSources.Item("PosDate").Value = Format(oRecordSet.Fields("PosDate").Value, "YYYYMMDD")
        
        
    oForm.Items("MSTCOD").Specific.VALUE = oRecordSet.Fields("MSTCOD").VALUE
    oForm.Items("FullName").Specific.VALUE = oRecordSet.Fields("FullName").VALUE
    
    oForm.DataSources.UserDataSources.Item("entno1").VALUE = oRecordSet.Fields("entno1").VALUE
    oForm.DataSources.UserDataSources.Item("servcomp1").VALUE = oRecordSet.Fields("servcomp1").VALUE
    oForm.DataSources.UserDataSources.Item("symd1").VALUE = Format(oRecordSet.Fields("symd1").VALUE, "YYYY-MM-DD")
    oForm.DataSources.UserDataSources.Item("eymd1").VALUE = Format(oRecordSet.Fields("eymd1").VALUE, "YYYY-MM-DD")

    oForm.DataSources.UserDataSources.Item("payrtot1").VALUE = oRecordSet.Fields("payrtot1").VALUE
    oForm.DataSources.UserDataSources.Item("bnstot1").VALUE = oRecordSet.Fields("bnstot1").VALUE
    
    oForm.DataSources.UserDataSources.Item("fwork1").VALUE = oRecordSet.Fields("fwork1").VALUE
    oForm.DataSources.UserDataSources.Item("ndtalw1").VALUE = oRecordSet.Fields("ndtalw1").VALUE
    oForm.DataSources.UserDataSources.Item("etcntax1").VALUE = oRecordSet.Fields("etcntax1").VALUE
    oForm.DataSources.UserDataSources.Item("lnchalw1").VALUE = oRecordSet.Fields("lnchalw1").VALUE
    oForm.DataSources.UserDataSources.Item("ftaxamt1").VALUE = oRecordSet.Fields("ftaxamt1").VALUE
    oForm.DataSources.UserDataSources.Item("savtaxddc1").VALUE = oRecordSet.Fields("savtaxddc1").VALUE
    oForm.DataSources.UserDataSources.Item("incmtax1").VALUE = oRecordSet.Fields("incmtax1").VALUE
    oForm.DataSources.UserDataSources.Item("residtax1").VALUE = oRecordSet.Fields("residtax1").VALUE
    oForm.DataSources.UserDataSources.Item("medcinsr1").VALUE = oRecordSet.Fields("medcinsr1").VALUE
    oForm.DataSources.UserDataSources.Item("asopinsr1").VALUE = oRecordSet.Fields("asopinsr1").VALUE
    oForm.DataSources.UserDataSources.Item("annuboamt1").VALUE = oRecordSet.Fields("annuboamt1").VALUE
    
    
    oForm.DataSources.UserDataSources.Item("entno2").VALUE = oRecordSet.Fields("entno2").VALUE
    oForm.DataSources.UserDataSources.Item("servcomp2").VALUE = oRecordSet.Fields("servcomp2").VALUE
    oForm.DataSources.UserDataSources.Item("symd2").VALUE = Format(oRecordSet.Fields("symd2").VALUE, "YYYY-MM-DD")
    oForm.DataSources.UserDataSources.Item("eymd2").VALUE = Format(oRecordSet.Fields("eymd2").VALUE, "YYYY-MM-DD")
    
    oForm.DataSources.UserDataSources.Item("payrtot2").VALUE = oRecordSet.Fields("payrtot2").VALUE
    oForm.DataSources.UserDataSources.Item("bnstot2").VALUE = oRecordSet.Fields("bnstot2").VALUE
    
    oForm.DataSources.UserDataSources.Item("fwork2").VALUE = oRecordSet.Fields("fwork2").VALUE
    oForm.DataSources.UserDataSources.Item("ndtalw2").VALUE = oRecordSet.Fields("ndtalw2").VALUE
    oForm.DataSources.UserDataSources.Item("etcntax2").VALUE = oRecordSet.Fields("etcntax2").VALUE
    oForm.DataSources.UserDataSources.Item("lnchalw2").VALUE = oRecordSet.Fields("lnchalw2").VALUE
    oForm.DataSources.UserDataSources.Item("ftaxamt2").VALUE = oRecordSet.Fields("ftaxamt2").VALUE
    oForm.DataSources.UserDataSources.Item("savtaxddc2").VALUE = oRecordSet.Fields("savtaxddc2").VALUE
    oForm.DataSources.UserDataSources.Item("indmtax2").VALUE = oRecordSet.Fields("indmtax2").VALUE
    oForm.DataSources.UserDataSources.Item("residtax2").VALUE = oRecordSet.Fields("residtax2").VALUE
    oForm.DataSources.UserDataSources.Item("medcinsr2").VALUE = oRecordSet.Fields("medcinsr2").VALUE
    oForm.DataSources.UserDataSources.Item("asopinsr2").VALUE = oRecordSet.Fields("asopinsr2").VALUE
    oForm.DataSources.UserDataSources.Item("annuboamt2").VALUE = oRecordSet.Fields("annuboamt2").VALUE
    
    
'    '//부서
    oForm.Items("TeamName").Specific.VALUE = oRecordSet.Fields("TeamName").VALUE
    oForm.Items("RspName").Specific.VALUE = oRecordSet.Fields("RspName").VALUE
    oForm.Items("ClsName").Specific.VALUE = oRecordSet.Fields("ClsName").VALUE
    
    
    
    oForm.Update
    
    
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Sub
PH_PY401_MTX02_Exit:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    
    Exit Sub
PH_PY401_MTX02_Error:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY401_MTX02_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY401_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY401_Validate_Error
    PH_PY401_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY401A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", bmt_Short, True
        PH_PY401_Validate = False
        GoTo PH_PY401_Validate_Exit
    End If
'
    If ValidateType = "수정" Then

    ElseIf ValidateType = "행삭제" Then

    ElseIf ValidateType = "취소" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY401_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY401_Validate_Error:
    PH_PY401_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY401_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

'//행삭제 (FormUID, pval, BubbleEvent, 매트릭스 이름, 디비데이터소스, 데이터 체크 필드명)
Private Sub Raise_EVENT_ROW_DELETE(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean, _
                                                    oMat As SAPbouiCOM.Matrix, DBData As SAPbouiCOM.DBDataSource, CheckField As String)

    Dim i As Long
    
    On Error GoTo Raise_EVENT_ROW_DELETE_Error

    If (oLastColRow > 0) Then
        If pval.BeforeAction = True Then
        
        ElseIf pval.BeforeAction = False Then
            If oMat.RowCount <> oMat.VisualRowCount Then
                oMat.FlushToDataSource
    
                While (i <= DBData.Size - 1)
                    If DBData.GetValue(CheckField, i) = "" Then
                        DBData.RemoveRecord (i)
                        i = 0
                    Else
                        i = i + 1
                    End If
                Wend
                
                For i = 0 To DBData.Size
                    Call DBData.setValue("U_LineNum", i, i + 1)
                Next i
                
                oMat.LoadFromDataSource
            End If
        End If
    End If
Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub PH_PY401_SAVE()

    '//데이타 저장
    
    Dim i       As Long
    Dim sQry    As String
    
    Dim CLTCOD$, MSTCOD$, FullName$, YEAR$
    
    
        
    Dim entno1$, servcomp1$, symd1$, eymd1$
    Dim entno2$, servcomp2$, symd2$, eymd2$
    
    Dim payrtot1, bnstot1, fwork1, ndtalw1, etcntax1, lnchalw1, ftaxamt1, savtaxddc1, incmtax1, fvsptax1, residtax1, medcinsr1, asopinsr1, annuboamt1 As Double
    Dim payrtot2, bnstot2, fwork2, ndtalw2, etcntax2, lnchalw2, ftaxamt2, savtaxddc2, indmtax2, fvsptax2, residtax2, medcinsr2, asopinsr2, annuboamt2 As Double
    
    entno1 = oForm.Items("entno1").Specific.VALUE
    servcomp1 = oForm.Items("servcomp1").Specific.VALUE
    symd1 = Trim(oForm.Items("symd1").Specific.VALUE)
    eymd1 = Trim(oForm.Items("eymd1").Specific.VALUE)
        
    payrtot1 = oForm.Items("payrtot1").Specific.VALUE
    bnstot1 = oForm.Items("bnstot1").Specific.VALUE
    fwork1 = oForm.Items("fwork1").Specific.VALUE
    ndtalw1 = oForm.Items("ndtalw1").Specific.VALUE
    etcntax1 = oForm.Items("etcntax1").Specific.VALUE
    lnchalw1 = oForm.Items("lnchalw1").Specific.VALUE
    ftaxamt1 = oForm.Items("ftaxamt1").Specific.VALUE
    savtaxddc1 = oForm.Items("savtaxddc1").Specific.VALUE
    incmtax1 = oForm.Items("incmtax1").Specific.VALUE
    fvsptax1 = oForm.Items("fvsptax1").Specific.VALUE
    residtax1 = oForm.Items("residtax1").Specific.VALUE
    medcinsr1 = oForm.Items("medcinsr1").Specific.VALUE
    asopinsr1 = oForm.Items("asopinsr1").Specific.VALUE
    annuboamt1 = oForm.Items("annuboamt1").Specific.VALUE
    
    entno2 = oForm.Items("entno2").Specific.VALUE
    servcomp2 = oForm.Items("servcomp2").Specific.VALUE
    If oForm.Items("symd2").Specific.VALUE = "" Then
        symd2 = ""
        eymd2 = ""
    Else
        symd2 = Trim(oForm.Items("symd2").Specific.VALUE)
        eymd2 = Trim(oForm.Items("eymd2").Specific.VALUE)
    End If
    

    payrtot2 = oForm.Items("payrtot2").Specific.VALUE
    bnstot2 = oForm.Items("bnstot2").Specific.VALUE
    fwork2 = oForm.Items("fwork2").Specific.VALUE
    ndtalw2 = oForm.Items("ndtalw2").Specific.VALUE
    etcntax2 = oForm.Items("etcntax2").Specific.VALUE
    lnchalw2 = oForm.Items("lnchalw2").Specific.VALUE
    ftaxamt2 = oForm.Items("ftaxamt2").Specific.VALUE
    savtaxddc2 = oForm.Items("savtaxddc2").Specific.VALUE
    indmtax2 = oForm.Items("indmtax2").Specific.VALUE
    fvsptax2 = oForm.Items("fvsptax2").Specific.VALUE
    residtax2 = oForm.Items("residtax2").Specific.VALUE
    medcinsr2 = oForm.Items("medcinsr2").Specific.VALUE
    asopinsr2 = oForm.Items("asopinsr2").Specific.VALUE
    annuboamt2 = oForm.Items("annuboamt2").Specific.VALUE
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY401_SAVE_Error

    Call oForm.Freeze(True)
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE)
    MSTCOD = oForm.Items("MSTCOD").Specific.VALUE
    FullName = oForm.Items("FullName").Specific.VALUE

    YEAR = oForm.Items("Year").Specific.VALUE
    
    If Trim(YEAR) = "" Then
        Call MDC_Com.MDC_GF_Message("년도가 없습니다. 확인바랍니다..", "E")
        GoTo PH_PY401_SAVE_Exit
    End If
    
    If Trim(CLTCOD) = "" Then
        Call MDC_Com.MDC_GF_Message("사업장이 없습니다. 확인바랍니다..", "E")
        GoTo PH_PY401_SAVE_Exit
    End If
    If Trim(MSTCOD) = "" Then
        Call MDC_Com.MDC_GF_Message("사번이 없습니다. 확인바랍니다..", "E")
        GoTo PH_PY401_SAVE_Exit
    End If
            
    sQry = " Select Count(*) From [p_sbservcomp] Where saup = '" & CLTCOD & "' And yyyy = '" & YEAR & "' And sabun = '" & MSTCOD & "'"
    oRecordSet.DoQuery sQry
    
    If oRecordSet.Fields(0).VALUE > 0 Then
        '//갱신
        
        sQry = "Update [p_sbservcomp] set "
        sQry = sQry + "entno1 = '" & entno1 & "',"
        sQry = sQry + "servcomp1 = '" & servcomp1 & "',"
        sQry = sQry + "symd1 = '" & symd1 & "',"
        sQry = sQry + "eymd1 = '" & eymd1 & "',"
        sQry = sQry + "payrtot1 = " & payrtot1 & ","
        sQry = sQry + "bnstot1 = " & bnstot1 & ","
        sQry = sQry + "fwork1 = " & fwork1 & ","
        sQry = sQry + "ndtalw1 = " & ndtalw1 & ","
        sQry = sQry + "etcntax1 = " & etcntax1 & ","
        sQry = sQry + "lnchalw1 = " & lnchalw1 & ","
        sQry = sQry + "ftaxamt1 = " & ftaxamt1 & ","
        sQry = sQry + "savtaxddc1 = " & savtaxddc1 & ","
        sQry = sQry + "incmtax1 = " & incmtax1 & ","
        sQry = sQry + "fvsptax1 = " & fvsptax1 & ","
        sQry = sQry + "residtax1 = " & residtax1 & ","
        sQry = sQry + "medcinsr1 = " & medcinsr1 & ","
        sQry = sQry + "asopinsr1 = " & asopinsr1 & ","
        sQry = sQry + "annuboamt1 =" & annuboamt1 & ","
        sQry = sQry + "entno2 = '" & entno2 & "',"
        sQry = sQry + "servcomp2 = '" & servcomp2 & "',"
        sQry = sQry + "symd2 = '" & symd2 & "',"
        sQry = sQry + "eymd2 = '" & eymd2 & "',"
        sQry = sQry + "payrtot2 = " & payrtot2 & ","
        sQry = sQry + "bnstot2= " & bnstot2 & ","
        sQry = sQry + "fwork2 = " & fwork2 & ","
        sQry = sQry + "ndtalw2 = " & ndtalw2 & ","
        sQry = sQry + "etcntax2 = " & etcntax2 & ","
        sQry = sQry + "lnchalw2 = " & lnchalw2 & ","
        sQry = sQry + "ftaxamt2 = " & ftaxamt2 & ","
        sQry = sQry + "savtaxddc2 = " & savtaxddc2 & ","
        sQry = sQry + "indmtax2 = " & indmtax2 & ","
        sQry = sQry + "fvsptax2 = " & fvsptax2 & ","
        sQry = sQry + "residtax2 = " & residtax2 & ","
        sQry = sQry + "medcinsr2 = " & medcinsr2 & ","
        sQry = sQry + "asopinsr2 = " & asopinsr2 & ","
        sQry = sQry + "annuboamt2 =" & annuboamt2
        
        sQry = sQry + " Where saup = '" & CLTCOD & "' And yyyy = '" & YEAR & "' And sabun = '" & MSTCOD & "'"
        
        oRecordSet.DoQuery sQry
        
    Else
        
        '//신규
        sQry = "INSERT INTO [p_sbservcomp]"
        sQry = sQry & " ("
        sQry = sQry & "saup,"
        sQry = sQry & "yyyy,"
        sQry = sQry & "sabun,"
        sQry = sQry & "entno1,"
        sQry = sQry & "servcomp1,"
        sQry = sQry & "symd1,"
        sQry = sQry & "eymd1,"
        sQry = sQry & "payrtot1,"
        sQry = sQry & "bnstot1,"
        sQry = sQry & "fwork1,"
        sQry = sQry & "ndtalw1,"
        sQry = sQry & "etcntax1,"
        sQry = sQry & "lnchalw1,"
        sQry = sQry & "ftaxamt1,"
        sQry = sQry & "savtaxddc1,"
        sQry = sQry & "incmtax1,"
        sQry = sQry & "fvsptax1,"
        sQry = sQry & "residtax1,"
        sQry = sQry & "medcinsr1,"
        sQry = sQry & "asopinsr1,"
        sQry = sQry & "annuboamt1,"
        sQry = sQry & "entno2,"
        sQry = sQry & "servcomp2,"
        sQry = sQry & "symd2,"
        sQry = sQry & "eymd2,"
        sQry = sQry & "payrtot2,"
        sQry = sQry & "bnstot2,"
        sQry = sQry & "fwork2,"
        sQry = sQry & "ndtalw2,"
        sQry = sQry & "etcntax2,"
        sQry = sQry & "lnchalw2,"
        sQry = sQry & "ftaxamt2,"
        sQry = sQry & "savtaxddc2,"
        sQry = sQry & "indmtax2,"
        sQry = sQry & "fvsptax2,"
        sQry = sQry & "residtax2,"
        sQry = sQry & "medcinsr2,"
        sQry = sQry & "asopinsr2,"
        sQry = sQry & "annuboamt2,"
        sQry = sQry & "jscntddc1,"
        sQry = sQry & "jscntddc2"
        sQry = sQry & " ) "
        sQry = sQry & "VALUES("

        sQry = sQry & "'" & CLTCOD & "',"
        sQry = sQry & "'" & YEAR & "',"
        sQry = sQry & "'" & MSTCOD & "',"
        sQry = sQry & "'" & entno1 & "',"
        sQry = sQry & "'" & servcomp1 & "',"
        sQry = sQry & "'" & symd1 & "',"
        sQry = sQry & "'" & eymd1 & "',"

        sQry = sQry & payrtot1 & ","
        sQry = sQry & bnstot1 & ","
        sQry = sQry & fwork1 & ","
        sQry = sQry & ndtalw1 & ","
        sQry = sQry & etcntax1 & ","
        sQry = sQry & lnchalw1 & ","
        sQry = sQry & ftaxamt1 & ","
        sQry = sQry & savtaxddc1 & ","
        sQry = sQry & incmtax1 & ","
        sQry = sQry & fvsptax1 & ","
        sQry = sQry & residtax1 & ","
        sQry = sQry & medcinsr1 & ","
        sQry = sQry & asopinsr1 & ","
        sQry = sQry & annuboamt1 & ","
        sQry = sQry & "'" & entno2 & "',"
        sQry = sQry & "'" & servcomp2 & "',"
        sQry = sQry & "'" & symd2 & "',"
        sQry = sQry & "'" & eymd2 & "',"
        sQry = sQry & payrtot2 & ","
        sQry = sQry & bnstot2 & ","
        sQry = sQry & fwork2 & ","
        sQry = sQry & ndtalw2 & ","
        sQry = sQry & etcntax2 & ","
        sQry = sQry & lnchalw2 & ","
        sQry = sQry & ftaxamt2 & ","
        sQry = sQry & savtaxddc2 & ","
        sQry = sQry & indmtax2 & ","
        sQry = sQry & fvsptax2 & ","
        sQry = sQry & residtax2 & ","
        sQry = sQry & medcinsr2 & ","
        sQry = sQry & asopinsr2 & ","
        sQry = sQry & annuboamt2 & ","
        sQry = sQry & "0, 0 )"

        oRecordSet.DoQuery sQry
    End If
    
    
    Call PH_PY401_FormItemEnabled


    Set oRecordSet = Nothing
    Call oForm.Freeze(False)

    Call PH_PY401_MTX01

    Exit Sub
    
PH_PY401_SAVE_Exit:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    
    Exit Sub
PH_PY401_SAVE_Error:
    Call oForm.Freeze(False)
   
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "PH_PY401_SAVE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub PH_PY401_Delete()
    '//선택된 자료 삭제
    
    Dim CLTCOD As String
    Dim MSTCOD As String
    Dim YEAR As String
    Dim FullName As String
    
    
    Dim i As Integer
    Dim cnt As Integer
    
    Dim sQry    As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY401_Delete_Error

    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    
    
    Call oForm.Freeze(True)
    
    CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE)
    YEAR = oForm.Items("Year").Specific.VALUE
    MSTCOD = oForm.Items("MSTCOD").Specific.VALUE
    FullName = oForm.Items("FullName").Specific.VALUE
    
    sQry = " Select Count(*) From [p_sbservcomp] Where saup = '" & CLTCOD & "' And yyyy = '" & YEAR & "' And sabun = '" & MSTCOD & "'"
    oRecordSet.DoQuery sQry
    
    cnt = oRecordSet.Fields(0).VALUE
    If cnt > 0 Then
    
        If Trim(YEAR) = "" Then
            Call MDC_Com.MDC_GF_Message("년도가 없습니다. 확인바랍니다..", "E")
            GoTo PH_PY401_Delete_Exit
        End If
        
        If Trim(CLTCOD) = "" Then
            Call MDC_Com.MDC_GF_Message("사업장이 없습니다. 확인바랍니다..", "E")
            GoTo PH_PY401_Delete_Exit
        End If
        If Trim(MSTCOD) = "" Then
            Call MDC_Com.MDC_GF_Message("사번이 없습니다. 확인바랍니다..", "E")
            GoTo PH_PY401_Delete_Exit
        End If
        
        
        
        
        If Sbo_Application.MessageBox(" 선택한사원('" & FullName & "')을 삭제하시겠습니까? ?", "2", "예", "아니오") = "1" Then
            sQry = "Delete From [p_sbservcomp] Where saup = '" & CLTCOD & "' AND  yyyy = '" & YEAR & "' And sabun = '" & MSTCOD & "' "
            oRecordSet.DoQuery sQry
        End If
    End If
    
   
    Call oForm.Freeze(False)
    
        
    Call PH_PY401_MTX01
    
    Set oRecordSet = Nothing
    
    
    Exit Sub
PH_PY401_Delete_Exit:
    Set oRecordSet = Nothing
        
    Call oForm.Freeze(False)
    Exit Sub
PH_PY401_Delete_Error:
    Set oRecordSet = Nothing
   
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY401_Delete_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub PH_PY401_TitleSetting(iRow As Long)
    Dim i               As Long
    Dim j               As Long
    Dim sQry            As String
     
    Dim COLNAM(5)       As String

    Dim oColumn         As SAPbouiCOM.EditTextColumn
    Dim oComboCol       As SAPbouiCOM.ComboBoxColumn
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
     
    On Error GoTo Error_Message
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    oForm.Freeze True
    
    COLNAM(0) = "년도"
    COLNAM(1) = "부서"
    COLNAM(2) = "담당"
    COLNAM(3) = "사번"
    COLNAM(4) = "성명"
    COLNAM(5) = "직급"

    For i = 0 To UBound(COLNAM)
        oGrid1.Columns.Item(i).TitleObject.Caption = COLNAM(i)
        oGrid1.Columns.Item(i).Editable = False
        
        oGrid1.Columns.Item(i).RightJustified = True
        
    Next i
    
    oGrid1.AutoResizeColumns

    oForm.Freeze False

    Set oColumn = Nothing

    Exit Sub

Error_Message:
    oForm.Freeze False
    Set oColumn = Nothing
    Sbo_Application.SetStatusBarMessage "PH_PY401_TitleSetting Error : " & Space(10) & Err.Description, bmt_Short, True
End Sub





