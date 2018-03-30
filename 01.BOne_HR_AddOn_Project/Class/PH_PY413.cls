VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY413"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY413.cls
'//  Module         : 인사관리 > 연말정산관리
'//  Desc           : 월세액.주택임차차입금자료 등록
'//********************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form

Public oGrid1           As SAPbouiCOM.Grid
Public oDS_PH_PY413A     As SAPbouiCOM.DataTable


Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long

Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY413.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY413_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY413"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
'    oForm.DataBrowser.BrowseBy = "Code"
        
    oForm.PaneLevel = 1
    oForm.Freeze True
    Call PH_PY413_CreateItems
    Call PH_PY413_FormItemEnabled
    Call PH_PY413_EnableMenus
'    Call PH_PY413_SetDocument(oFromDocEntry01)
'    Call PH_PY413_FormResize
    
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

Private Function PH_PY413_CreateItems() As Boolean

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
    
    On Error GoTo PH_PY413_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    'Set oGrid1 = oForm.Items("Grid01").Specific
    
    oForm.DataSources.DataTables.Add ("PH_PY413")
    
    'oGrid1.DataTable = oForm.DataSources.DataTables.Item("PH_PY413")
    Set oDS_PH_PY413A = oForm.DataSources.DataTables.Item("PH_PY413")
    

    '//----------------------------------------------------------------------------------------------
    '// 기본사항
    '//----------------------------------------------------------------------------------------------
    
    '//사업장
    
    Set oCombo = oForm.Items("CLTCOD").Specific
'    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
'    Call SetReDataCombo(oForm, sQry, oCombo)
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
    
    '//임대인성명
    Call oForm.DataSources.UserDataSources.Add("ws_name1", dt_SHORT_TEXT, 20)
    oForm.Items("ws_name1").Specific.DataBind.SetBound True, "", "ws_name1"
    
    Call oForm.DataSources.UserDataSources.Add("ws_name2", dt_SHORT_TEXT, 20)
    oForm.Items("ws_name2").Specific.DataBind.SetBound True, "", "ws_name2"
    
    Call oForm.DataSources.UserDataSources.Add("ws_name3", dt_SHORT_TEXT, 20)
    oForm.Items("ws_name3").Specific.DataBind.SetBound True, "", "ws_name3"
    
    '//임대인주민등록번호
    Call oForm.DataSources.UserDataSources.Add("ws_jumin1", dt_SHORT_TEXT, 13)
    oForm.Items("ws_jumin1").Specific.DataBind.SetBound True, "", "ws_jumin1"
    
    Call oForm.DataSources.UserDataSources.Add("ws_jumin2", dt_SHORT_TEXT, 13)
    oForm.Items("ws_jumin2").Specific.DataBind.SetBound True, "", "ws_jumin2"
    
    Call oForm.DataSources.UserDataSources.Add("ws_jumin3", dt_SHORT_TEXT, 13)
    oForm.Items("ws_jumin3").Specific.DataBind.SetBound True, "", "ws_jumin3"
    
    '//주택유형
    Set oCombo = oForm.Items("ws_hcode1").Specific
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '80' AND U_UseYN= 'Y'"
    Call SetReDataCombo(oForm, sQry, oCombo, "Y")
    oCombo.Select 0, psk_Index
    oForm.Items("ws_hcode1").DisplayDesc = True
    
    Set oCombo = oForm.Items("ws_hcode2").Specific
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '80' AND U_UseYN= 'Y'"
    Call SetReDataCombo(oForm, sQry, oCombo, "Y")
    oCombo.Select 0, psk_Index
    oForm.Items("ws_hcode2").DisplayDesc = True
    
    Set oCombo = oForm.Items("ws_hcode3").Specific
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '80' AND U_UseYN= 'Y'"
    Call SetReDataCombo(oForm, sQry, oCombo, "Y")
    oCombo.Select 0, psk_Index
    oForm.Items("ws_hcode3").DisplayDesc = True
    
    '//주택계약면적
    Call oForm.DataSources.UserDataSources.Add("ws_hm1", dt_QUANTITY)
    oForm.Items("ws_hm1").Specific.DataBind.SetBound True, "", "ws_hm1"
    
    Call oForm.DataSources.UserDataSources.Add("ws_hm2", dt_QUANTITY)
    oForm.Items("ws_hm2").Specific.DataBind.SetBound True, "", "ws_hm2"
    
    Call oForm.DataSources.UserDataSources.Add("ws_hm3", dt_QUANTITY)
    oForm.Items("ws_hm3").Specific.DataBind.SetBound True, "", "ws_hm3"
    
    '//임대차계약서상주소
    Call oForm.DataSources.UserDataSources.Add("ws_addr1", dt_SHORT_TEXT, 100)
    oForm.Items("ws_addr1").Specific.DataBind.SetBound True, "", "ws_addr1"
    
    Call oForm.DataSources.UserDataSources.Add("ws_addr2", dt_SHORT_TEXT, 100)
    oForm.Items("ws_addr2").Specific.DataBind.SetBound True, "", "ws_addr2"
    
    Call oForm.DataSources.UserDataSources.Add("ws_addr3", dt_SHORT_TEXT, 100)
    oForm.Items("ws_addr3").Specific.DataBind.SetBound True, "", "ws_addr3"
    
    '//임대차계약기간
    Call oForm.DataSources.UserDataSources.Add("ws_fymd1", dt_DATE, 8)
    oForm.Items("ws_fymd1").Specific.DataBind.SetBound True, "", "ws_fymd1"
    
    Call oForm.DataSources.UserDataSources.Add("ws_fymd2", dt_DATE, 8)
    oForm.Items("ws_fymd2").Specific.DataBind.SetBound True, "", "ws_fymd2"
    
    Call oForm.DataSources.UserDataSources.Add("ws_fymd3", dt_DATE, 8)
    oForm.Items("ws_fymd3").Specific.DataBind.SetBound True, "", "ws_fymd3"
    
    Call oForm.DataSources.UserDataSources.Add("ws_tymd1", dt_DATE, 8)
    oForm.Items("ws_tymd1").Specific.DataBind.SetBound True, "", "ws_tymd1"
    
    Call oForm.DataSources.UserDataSources.Add("ws_tymd2", dt_DATE, 8)
    oForm.Items("ws_tymd2").Specific.DataBind.SetBound True, "", "ws_tymd2"
    
    Call oForm.DataSources.UserDataSources.Add("ws_tymd3", dt_DATE, 8)
    oForm.Items("ws_tymd3").Specific.DataBind.SetBound True, "", "ws_tymd3"
    
    '//월세액
    Call oForm.DataSources.UserDataSources.Add("ws_mamt1", dt_SUM)
    oForm.Items("ws_mamt1").Specific.DataBind.SetBound True, "", "ws_mamt1"
    
    Call oForm.DataSources.UserDataSources.Add("ws_mamt2", dt_SUM)
    oForm.Items("ws_mamt2").Specific.DataBind.SetBound True, "", "ws_mamt2"
    
    Call oForm.DataSources.UserDataSources.Add("ws_mamt3", dt_SUM)
    oForm.Items("ws_mamt3").Specific.DataBind.SetBound True, "", "ws_mamt3"
    
    '//공제금액
    Call oForm.DataSources.UserDataSources.Add("ws_gamt1", dt_SUM)
    oForm.Items("ws_gamt1").Specific.DataBind.SetBound True, "", "ws_gamt1"
    
    Call oForm.DataSources.UserDataSources.Add("ws_gamt2", dt_SUM)
    oForm.Items("ws_gamt2").Specific.DataBind.SetBound True, "", "ws_gamt2"
    
    Call oForm.DataSources.UserDataSources.Add("ws_gamt3", dt_SUM)
    oForm.Items("ws_gamt3").Specific.DataBind.SetBound True, "", "ws_gamt3"
    
    '//대주
    Call oForm.DataSources.UserDataSources.Add("dj_name1", dt_SHORT_TEXT, 20)
    oForm.Items("dj_name1").Specific.DataBind.SetBound True, "", "dj_name1"
    
    Call oForm.DataSources.UserDataSources.Add("dj_name2", dt_SHORT_TEXT, 20)
    oForm.Items("dj_name2").Specific.DataBind.SetBound True, "", "dj_name2"
    
    Call oForm.DataSources.UserDataSources.Add("dj_name3", dt_SHORT_TEXT, 20)
    oForm.Items("dj_name3").Specific.DataBind.SetBound True, "", "dj_name3"
    
    '//대주 주민등록번호
    Call oForm.DataSources.UserDataSources.Add("dj_jumin1", dt_SHORT_TEXT, 13)
    oForm.Items("dj_jumin1").Specific.DataBind.SetBound True, "", "dj_jumin1"
    
    Call oForm.DataSources.UserDataSources.Add("dj_jumin2", dt_SHORT_TEXT, 13)
    oForm.Items("dj_jumin2").Specific.DataBind.SetBound True, "", "dj_jumin2"
    
    Call oForm.DataSources.UserDataSources.Add("dj_jumin3", dt_SHORT_TEXT, 13)
    oForm.Items("dj_jumin3").Specific.DataBind.SetBound True, "", "dj_jumin3"
    
    '//금전소비대차계약기간
    Call oForm.DataSources.UserDataSources.Add("dj_fymd1", dt_DATE, 8)
    oForm.Items("dj_fymd1").Specific.DataBind.SetBound True, "", "dj_fymd1"
    
    Call oForm.DataSources.UserDataSources.Add("dj_fymd2", dt_DATE, 8)
    oForm.Items("dj_fymd2").Specific.DataBind.SetBound True, "", "dj_fymd2"
    
    Call oForm.DataSources.UserDataSources.Add("dj_fymd3", dt_DATE, 8)
    oForm.Items("dj_fymd3").Specific.DataBind.SetBound True, "", "dj_fymd3"
    
    Call oForm.DataSources.UserDataSources.Add("dj_tymd1", dt_DATE, 8)
    oForm.Items("dj_tymd1").Specific.DataBind.SetBound True, "", "dj_tymd1"
    
    Call oForm.DataSources.UserDataSources.Add("dj_tymd2", dt_DATE, 8)
    oForm.Items("dj_tymd2").Specific.DataBind.SetBound True, "", "dj_tymd2"
    
    Call oForm.DataSources.UserDataSources.Add("dj_tymd3", dt_DATE, 8)
    oForm.Items("dj_tymd3").Specific.DataBind.SetBound True, "", "dj_tymd3"
    
    '//차임급이자율
    Call oForm.DataSources.UserDataSources.Add("dj_eja1", dt_SHORT_TEXT, 4)
    oForm.Items("dj_eja1").Specific.DataBind.SetBound True, "", "dj_eja1"
    
    Call oForm.DataSources.UserDataSources.Add("dj_eja2", dt_SHORT_TEXT, 4)
    oForm.Items("dj_eja2").Specific.DataBind.SetBound True, "", "dj_eja2"
    
    Call oForm.DataSources.UserDataSources.Add("dj_eja3", dt_SHORT_TEXT, 4)
    oForm.Items("dj_eja3").Specific.DataBind.SetBound True, "", "dj_eja3"
    
    '//계
    Call oForm.DataSources.UserDataSources.Add("dj_tamt1", dt_SUM)
    oForm.Items("dj_tamt1").Specific.DataBind.SetBound True, "", "dj_tamt1"
    
    Call oForm.DataSources.UserDataSources.Add("dj_tamt2", dt_SUM)
    oForm.Items("dj_tamt2").Specific.DataBind.SetBound True, "", "dj_tamt2"
    
    Call oForm.DataSources.UserDataSources.Add("dj_tamt3", dt_SUM)
    oForm.Items("dj_tamt3").Specific.DataBind.SetBound True, "", "dj_tamt3"
    
    '//원리금
    Call oForm.DataSources.UserDataSources.Add("dj_wamt1", dt_SUM)
    oForm.Items("dj_wamt1").Specific.DataBind.SetBound True, "", "dj_wamt1"
    
    Call oForm.DataSources.UserDataSources.Add("dj_wamt2", dt_SUM)
    oForm.Items("dj_wamt2").Specific.DataBind.SetBound True, "", "dj_wamt2"
    
    Call oForm.DataSources.UserDataSources.Add("dj_wamt3", dt_SUM)
    oForm.Items("dj_wamt3").Specific.DataBind.SetBound True, "", "dj_wamt3"
    
    '//이자
    Call oForm.DataSources.UserDataSources.Add("dj_eamt1", dt_SUM)
    oForm.Items("dj_eamt1").Specific.DataBind.SetBound True, "", "dj_eamt1"
    
    Call oForm.DataSources.UserDataSources.Add("dj_eamt2", dt_SUM)
    oForm.Items("dj_eamt2").Specific.DataBind.SetBound True, "", "dj_eamt2"
    
    Call oForm.DataSources.UserDataSources.Add("dj_eamt3", dt_SUM)
    oForm.Items("dj_eamt3").Specific.DataBind.SetBound True, "", "dj_eamt3"
    
    '//공제금액
    Call oForm.DataSources.UserDataSources.Add("dj_gamt1", dt_SUM)
    oForm.Items("dj_gamt1").Specific.DataBind.SetBound True, "", "dj_gamt1"
    
    Call oForm.DataSources.UserDataSources.Add("dj_gamt2", dt_SUM)
    oForm.Items("dj_gamt2").Specific.DataBind.SetBound True, "", "dj_gamt2"
    
    Call oForm.DataSources.UserDataSources.Add("dj_gamt3", dt_SUM)
    oForm.Items("dj_gamt3").Specific.DataBind.SetBound True, "", "dj_gamt3"
    
    '//임대인성명
    Call oForm.DataSources.UserDataSources.Add("ld_name1", dt_SHORT_TEXT, 20)
    oForm.Items("ld_name1").Specific.DataBind.SetBound True, "", "ld_name1"
    
    Call oForm.DataSources.UserDataSources.Add("ld_name2", dt_SHORT_TEXT, 20)
    oForm.Items("ld_name2").Specific.DataBind.SetBound True, "", "ld_name2"
    
    Call oForm.DataSources.UserDataSources.Add("ld_name3", dt_SHORT_TEXT, 20)
    oForm.Items("ld_name3").Specific.DataBind.SetBound True, "", "ld_name3"
    
    '//임대인주민등록번호
    Call oForm.DataSources.UserDataSources.Add("ld_jumin1", dt_SHORT_TEXT, 13)
    oForm.Items("ld_jumin1").Specific.DataBind.SetBound True, "", "ld_jumin1"
    
    Call oForm.DataSources.UserDataSources.Add("ld_jumin2", dt_SHORT_TEXT, 13)
    oForm.Items("ld_jumin2").Specific.DataBind.SetBound True, "", "ld_jumin2"
    
    Call oForm.DataSources.UserDataSources.Add("ld_jumin3", dt_SHORT_TEXT, 13)
    oForm.Items("ld_jumin3").Specific.DataBind.SetBound True, "", "ld_jumin3"
    
    '//주택유형
    Set oCombo = oForm.Items("ld_hcode1").Specific
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '80' AND U_UseYN= 'Y'"
    Call SetReDataCombo(oForm, sQry, oCombo, "Y")
    oCombo.Select 0, psk_Index
    oForm.Items("ld_hcode1").DisplayDesc = True
    
    Set oCombo = oForm.Items("ld_hcode2").Specific
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '80' AND U_UseYN= 'Y'"
    Call SetReDataCombo(oForm, sQry, oCombo, "Y")
    oCombo.Select 0, psk_Index
    oForm.Items("ld_hcode2").DisplayDesc = True
    
    Set oCombo = oForm.Items("ld_hcode3").Specific
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '80' AND U_UseYN= 'Y'"
    Call SetReDataCombo(oForm, sQry, oCombo, "Y")
    oCombo.Select 0, psk_Index
    oForm.Items("ld_hcode3").DisplayDesc = True
    
    '//주택계약면적
    Call oForm.DataSources.UserDataSources.Add("ld_hm1", dt_QUANTITY)
    oForm.Items("ld_hm1").Specific.DataBind.SetBound True, "", "ld_hm1"
    
    Call oForm.DataSources.UserDataSources.Add("ld_hm2", dt_QUANTITY)
    oForm.Items("ld_hm2").Specific.DataBind.SetBound True, "", "ld_hm2"
    
    Call oForm.DataSources.UserDataSources.Add("ld_hm3", dt_QUANTITY)
    oForm.Items("ld_hm3").Specific.DataBind.SetBound True, "", "ld_hm3"
    
    '//임대차계약서상주소
    Call oForm.DataSources.UserDataSources.Add("ld_addr1", dt_SHORT_TEXT, 100)
    oForm.Items("ld_addr1").Specific.DataBind.SetBound True, "", "ld_addr1"
    
    Call oForm.DataSources.UserDataSources.Add("ld_addr2", dt_SHORT_TEXT, 100)
    oForm.Items("ld_addr2").Specific.DataBind.SetBound True, "", "ld_addr2"
    
    Call oForm.DataSources.UserDataSources.Add("ld_addr3", dt_SHORT_TEXT, 100)
    oForm.Items("ld_addr3").Specific.DataBind.SetBound True, "", "ld_addr3"
    
    '//임대차계약기간
    Call oForm.DataSources.UserDataSources.Add("ld_fymd1", dt_DATE, 8)
    oForm.Items("ld_fymd1").Specific.DataBind.SetBound True, "", "ld_fymd1"
    
    Call oForm.DataSources.UserDataSources.Add("ld_fymd2", dt_DATE, 8)
    oForm.Items("ld_fymd2").Specific.DataBind.SetBound True, "", "ld_fymd2"
    
    Call oForm.DataSources.UserDataSources.Add("ld_fymd3", dt_DATE, 8)
    oForm.Items("ld_fymd3").Specific.DataBind.SetBound True, "", "ld_fymd3"
    
    Call oForm.DataSources.UserDataSources.Add("ld_tymd1", dt_DATE, 8)
    oForm.Items("ld_tymd1").Specific.DataBind.SetBound True, "", "ld_tymd1"
    
    Call oForm.DataSources.UserDataSources.Add("ld_tymd2", dt_DATE, 8)
    oForm.Items("ld_tymd2").Specific.DataBind.SetBound True, "", "ld_tymd2"
    
    Call oForm.DataSources.UserDataSources.Add("ld_tymd3", dt_DATE, 8)
    oForm.Items("ld_tymd3").Specific.DataBind.SetBound True, "", "ld_tymd3"
    
    '//전세보증금
    Call oForm.DataSources.UserDataSources.Add("ld_bamt1", dt_SUM)
    oForm.Items("ld_bamt1").Specific.DataBind.SetBound True, "", "ld_bamt1"
    
    Call oForm.DataSources.UserDataSources.Add("ld_bamt2", dt_SUM)
    oForm.Items("ld_bamt2").Specific.DataBind.SetBound True, "", "ld_bamt2"
    
    Call oForm.DataSources.UserDataSources.Add("ld_bamt3", dt_SUM)
    oForm.Items("ld_bamt3").Specific.DataBind.SetBound True, "", "ld_bamt3"
    
    
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
    
PH_PY413_CreateItems_Error:
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY413_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY413_EnableMenus()
    
    On Error GoTo PH_PY413_EnableMenus_Error

    Call oForm.EnableMenu("1283", False)     '//제거
    Call oForm.EnableMenu("1284", False)    '//취소
    Call oForm.EnableMenu("1293", False)     '//행삭제
    
    Exit Sub
    
PH_PY413_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY413_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY413_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY413_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY413_FormItemEnabled
'        Call PH_PY413_AddMatrixRow
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY413_FormItemEnabled
        oForm.Items("Code").Specific.VALUE = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY413_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY413_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY413_FormItemEnabled()
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim sQry        As String
    Dim i           As Long
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Dim CLTCOD      As String
    Dim sPosDate    As String
    
       
    On Error GoTo PH_PY413_FormItemEnabled_Error
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Call oForm.Freeze(True)
    If (oForm.Mode = fm_ADD_MODE) Then
    
        Call oForm.EnableMenu("1281", False)     '//문서찾기
        Call oForm.EnableMenu("1282", True)      '//문서추가
        
        oForm.Items("Year").Specific.VALUE = Format(Now, "YYYY") - 1
        'oForm.Items("MSTCOD").Specific.VALUE = ""
        'oForm.Items("FullName").Specific.VALUE = ""
        'oForm.Items("TeamName").Specific.VALUE = ""
        'oForm.Items("RspName").Specific.VALUE = ""
        'oForm.Items("ClsName").Specific.VALUE = ""
    '-----------------------------------------------
    oForm.DataSources.UserDataSources.Item("ws_name1").VALUE = ""
    oForm.DataSources.UserDataSources.Item("ws_name2").VALUE = ""
    oForm.DataSources.UserDataSources.Item("ws_name3").VALUE = ""
    oForm.DataSources.UserDataSources.Item("ws_jumin1").VALUE = ""
    oForm.DataSources.UserDataSources.Item("ws_jumin2").VALUE = ""
    oForm.DataSources.UserDataSources.Item("ws_jumin3").VALUE = ""
    Set oCombo = oForm.Items("ws_hcode1").Specific
        oCombo.Select 0, psk_Index
    Set oCombo = oForm.Items("ws_hcode2").Specific
        oCombo.Select 0, psk_Index
    Set oCombo = oForm.Items("ws_hcode3").Specific
        oCombo.Select 0, psk_Index
    oForm.DataSources.UserDataSources.Item("ws_hm1").VALUE = 0
    oForm.DataSources.UserDataSources.Item("ws_hm2").VALUE = 0
    oForm.DataSources.UserDataSources.Item("ws_hm3").VALUE = 0
    oForm.DataSources.UserDataSources.Item("ws_addr1").VALUE = ""
    oForm.DataSources.UserDataSources.Item("ws_addr2").VALUE = ""
    oForm.DataSources.UserDataSources.Item("ws_addr3").VALUE = ""
    oForm.DataSources.UserDataSources.Item("ws_fymd1").VALUE = ""
    oForm.DataSources.UserDataSources.Item("ws_fymd2").VALUE = ""
    oForm.DataSources.UserDataSources.Item("ws_fymd3").VALUE = ""
    oForm.DataSources.UserDataSources.Item("ws_tymd1").VALUE = ""
    oForm.DataSources.UserDataSources.Item("ws_tymd2").VALUE = ""
    oForm.DataSources.UserDataSources.Item("ws_tymd3").VALUE = ""
    oForm.DataSources.UserDataSources.Item("ws_mamt1").VALUE = 0
    oForm.DataSources.UserDataSources.Item("ws_mamt2").VALUE = 0
    oForm.DataSources.UserDataSources.Item("ws_mamt3").VALUE = 0
    oForm.DataSources.UserDataSources.Item("ws_gamt1").VALUE = 0
    oForm.DataSources.UserDataSources.Item("ws_gamt2").VALUE = 0
    oForm.DataSources.UserDataSources.Item("ws_gamt3").VALUE = 0
    
    oForm.DataSources.UserDataSources.Item("dj_name1").VALUE = ""
    oForm.DataSources.UserDataSources.Item("dj_name2").VALUE = ""
    oForm.DataSources.UserDataSources.Item("dj_name3").VALUE = ""
    oForm.DataSources.UserDataSources.Item("dj_jumin1").VALUE = ""
    oForm.DataSources.UserDataSources.Item("dj_jumin2").VALUE = ""
    oForm.DataSources.UserDataSources.Item("dj_jumin3").VALUE = ""
    oForm.DataSources.UserDataSources.Item("dj_fymd1").VALUE = ""
    oForm.DataSources.UserDataSources.Item("dj_fymd2").VALUE = ""
    oForm.DataSources.UserDataSources.Item("dj_fymd3").VALUE = ""
    oForm.DataSources.UserDataSources.Item("dj_tymd1").VALUE = ""
    oForm.DataSources.UserDataSources.Item("dj_tymd2").VALUE = ""
    oForm.DataSources.UserDataSources.Item("dj_tymd3").VALUE = ""
    oForm.DataSources.UserDataSources.Item("dj_eja1").VALUE = ""
    oForm.DataSources.UserDataSources.Item("dj_eja2").VALUE = ""
    oForm.DataSources.UserDataSources.Item("dj_eja3").VALUE = ""
    oForm.DataSources.UserDataSources.Item("dj_tamt1").VALUE = 0
    oForm.DataSources.UserDataSources.Item("dj_tamt2").VALUE = 0
    oForm.DataSources.UserDataSources.Item("dj_tamt3").VALUE = 0
    oForm.DataSources.UserDataSources.Item("dj_wamt1").VALUE = 0
    oForm.DataSources.UserDataSources.Item("dj_wamt2").VALUE = 0
    oForm.DataSources.UserDataSources.Item("dj_wamt3").VALUE = 0
    oForm.DataSources.UserDataSources.Item("dj_eamt1").VALUE = 0
    oForm.DataSources.UserDataSources.Item("dj_eamt2").VALUE = 0
    oForm.DataSources.UserDataSources.Item("dj_eamt3").VALUE = 0
    oForm.DataSources.UserDataSources.Item("dj_gamt1").VALUE = 0
    oForm.DataSources.UserDataSources.Item("dj_gamt2").VALUE = 0
    oForm.DataSources.UserDataSources.Item("dj_gamt3").VALUE = 0
    
    oForm.DataSources.UserDataSources.Item("ld_name1").VALUE = ""
    oForm.DataSources.UserDataSources.Item("ld_name2").VALUE = ""
    oForm.DataSources.UserDataSources.Item("ld_name3").VALUE = ""
    oForm.DataSources.UserDataSources.Item("ld_jumin1").VALUE = ""
    oForm.DataSources.UserDataSources.Item("ld_jumin2").VALUE = ""
    oForm.DataSources.UserDataSources.Item("ld_jumin3").VALUE = ""
    Set oCombo = oForm.Items("ld_hcode1").Specific
        oCombo.Select 0, psk_Index
    Set oCombo = oForm.Items("ld_hcode2").Specific
        oCombo.Select 0, psk_Index
    Set oCombo = oForm.Items("ld_hcode3").Specific
        oCombo.Select 0, psk_Index
    oForm.DataSources.UserDataSources.Item("ld_hm1").VALUE = 0
    oForm.DataSources.UserDataSources.Item("ld_hm2").VALUE = 0
    oForm.DataSources.UserDataSources.Item("ld_hm3").VALUE = 0
    oForm.DataSources.UserDataSources.Item("ld_addr1").VALUE = ""
    oForm.DataSources.UserDataSources.Item("ld_addr2").VALUE = ""
    oForm.DataSources.UserDataSources.Item("ld_addr3").VALUE = ""
    oForm.DataSources.UserDataSources.Item("ld_fymd1").VALUE = ""
    oForm.DataSources.UserDataSources.Item("ld_fymd2").VALUE = ""
    oForm.DataSources.UserDataSources.Item("ld_fymd3").VALUE = ""
    oForm.DataSources.UserDataSources.Item("ld_tymd1").VALUE = ""
    oForm.DataSources.UserDataSources.Item("ld_tymd2").VALUE = ""
    oForm.DataSources.UserDataSources.Item("ld_tymd3").VALUE = ""
    oForm.DataSources.UserDataSources.Item("ld_bamt1").VALUE = 0
    oForm.DataSources.UserDataSources.Item("ld_bamt2").VALUE = 0
    oForm.DataSources.UserDataSources.Item("ld_bamt3").VALUE = 0
  
   '-----------------------------------------------
   
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
    
    '//Key set
    oForm.Items("CLTCOD").Enabled = True
    oForm.Items("Year").Enabled = True
    oForm.Items("MSTCOD").Enabled = True
    
    Set oCombo = Nothing
    Call oForm.Freeze(False)
    Exit Sub
    
PH_PY413_FormItemEnabled_Error:
    Set oCombo = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY413_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
    Dim FullName    As String
    Dim Amt         As Long
    Dim gamt        As Long
    Dim YY          As String
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
                    If PH_PY413_DataValidCheck = False Then
                        BubbleEvent = False
                    End If
                End If
                
                If pval.ItemUID = "Btn_ret" Then
                    Call PH_PY413_MTX01
                End If
                
                If pval.ItemUID = "Btn01" Then
                
                    CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE)
                    YY = oForm.Items("Year").Specific.VALUE
                    sQry = " Select CODE from [@PSH_PAYBACK] WHERE CODE = '" & YY & CLTCOD & "'"
                    oRecordSet.DoQuery sQry
                           
                        Result = oRecordSet.Fields(0).VALUE
                    If Result <> "" Then
                        Call MDC_Com.MDC_GF_Message("등록불가 년도입니다. 담당자에게 문의바랍니다.", "E")
                    End If
                    If Result = "" Then
                        Call PH_PY413_SAVE
                    End If
                End If
                
                                
                If pval.ItemUID = "Btn_del" Then
                    Call PH_PY413_Delete
                    Call PH_PY413_FormItemEnabled
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
                            Call PH_PY413_FormItemEnabled
                        End If
                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY413_FormItemEnabled
                        End If
                    ElseIf oForm.Mode = fm_OK_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY413_FormItemEnabled
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
                            PH_PY413_MTX02 pval.ItemUID, pval.Row, pval.ColUID
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
                            
                            oForm.DataSources.UserDataSources.Item("FullName").VALUE = oRecordSet.Fields("FullName").VALUE
'                            oForm.Items("FullName").Specific.VALUE = oRecordSet.Fields("FullName").VALUE
                            oForm.Items("TeamName").Specific.VALUE = oRecordSet.Fields("TeamName").VALUE
                            oForm.Items("RspName").Specific.VALUE = oRecordSet.Fields("RspName").VALUE
                            oForm.Items("ClsName").Specific.VALUE = oRecordSet.Fields("ClsName").VALUE
                            
                            '//조회실행
                            Call PH_PY413_MTX01
                        Case "FullName"
                            CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE)
                            FullName = oForm.Items("FullName").Specific.VALUE
                            
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
                            sQry = sQry + " And U_status <> '5'"      '퇴사자 제외
                            sQry = sQry + " and U_FullName = '" + FullName + "'"
                              
                            oRecordSet.DoQuery sQry
                            
                            oForm.DataSources.UserDataSources.Item("MSTCOD").VALUE = oRecordSet.Fields("Code").VALUE
'                            oForm.Items("MSTCOD").Specific.VALUE = oRecordSet.Fields("Code").VALUE
                            oForm.Items("TeamName").Specific.VALUE = oRecordSet.Fields("TeamName").VALUE
                            oForm.Items("RspName").Specific.VALUE = oRecordSet.Fields("RspName").VALUE
                            oForm.Items("ClsName").Specific.VALUE = oRecordSet.Fields("ClsName").VALUE
                            
                            '//조회실행
                            Call PH_PY413_MTX01
                            
                        Case "ws_mamt1"
                            Amt = 0
                            gamt = 0
                            Amt = Round(oForm.Items("ws_mamt1").Specific.VALUE * 0.1, 0)
                            gamt = Amt
                            If gamt > 750000 Then
                               oForm.Items("ws_gamt1").Specific.VALUE = 750000  '14년한도7,500,000 에 세액공제10%(75만원)
                            Else
                               oForm.Items("ws_gamt1").Specific.VALUE = gamt
                            End If
                                    
                            If oForm.Items("ws_gamt1").Specific.VALUE < 0 Then
                               oForm.Items("ws_gamt1").Specific.VALUE = 0
                            End If
                            
                        Case "ws_mamt2"
                            Amt = 0
                            gamt = 0
                            Amt = Round(oForm.Items("ws_mamt2").Specific.VALUE * 0.1, 0)
                            gamt = oForm.Items("ws_gamt1").Specific.VALUE
                            
                            If gamt + Amt > 750000 Then
                               oForm.Items("ws_gamt2").Specific.VALUE = 750000 - gamt
                            Else
                               oForm.Items("ws_gamt2").Specific.VALUE = Amt
                            End If
                                    
                            If oForm.Items("ws_gamt2").Specific.VALUE < 0 Then
                               oForm.Items("ws_gamt2").Specific.VALUE = 0
                            End If
                            
                         Case "ws_mamt3"
                            Amt = 0
                            gamt = 0
                            Amt = Round(oForm.Items("ws_mamt3").Specific.VALUE * 0.1, 0)
                            gamt = oForm.Items("ws_gamt1").Specific.VALUE + oForm.Items("ws_gamt2").Specific.VALUE
                            If gamt + Amt > 750000 Then
                               oForm.Items("ws_gamt3").Specific.VALUE = 750000 - gamt
                            Else
                               oForm.Items("ws_gamt3").Specific.VALUE = Amt
                            End If
                                    
                            If oForm.Items("ws_gamt3").Specific.VALUE < 0 Then
                               oForm.Items("ws_gamt3").Specific.VALUE = 0
                            End If
                            
                         Case "dj_tamt1"
                            Amt = 0
                            gamt = 0
                            Amt = oForm.Items("dj_tamt1").Specific.VALUE
                            gamt = Round(Amt * 0.4, 0)
                            If gamt > 3000000 Then
                               oForm.Items("dj_gamt1").Specific.VALUE = 3000000
                            Else
                               oForm.Items("dj_gamt1").Specific.VALUE = gamt
                            End If
                                    
                            If oForm.Items("dj_gamt1").Specific.VALUE < 0 Then
                               oForm.Items("dj_gamt1").Specific.VALUE = 0
                            End If
                            
                        Case "dj_tamt2"
                            Amt = 0
                            gamt = 0
                            Amt = oForm.Items("dj_tamt2").Specific.VALUE
                            gamt = Round(Amt * 0.4, 0)
                            If gamt > 3000000 Then
                               oForm.Items("dj_gamt2").Specific.VALUE = 3000000
                            Else
                               oForm.Items("dj_gamt2").Specific.VALUE = gamt
                            End If
                                    
                            If oForm.Items("dj_gamt2").Specific.VALUE < 0 Then
                               oForm.Items("dj_gamt2").Specific.VALUE = 0
                            End If
                            
                        Case "dj_tamt3"
                            Amt = 0
                            gamt = 0
                            Amt = oForm.Items("dj_tamt3").Specific.VALUE
                            gamt = Round(Amt * 0.4, 0)
                            If gamt > 3000000 Then
                               oForm.Items("dj_gamt3").Specific.VALUE = 3000000
                            Else
                               oForm.Items("dj_gamt3").Specific.VALUE = gamt
                            End If
                                    
                            If oForm.Items("dj_gamt3").Specific.VALUE < 0 Then
                               oForm.Items("dj_gamt3").Specific.VALUE = 0
                            End If
                        
                    End Select
                 
                End If
            End If
'            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_MATRIX_LOAD: '//11
            If pval.BeforeAction = True Then
            ElseIf pval.BeforeAction = False Then
'                oMat1.LoadFromDataSource
'                Call PH_PY413_AddMatrixRow
                
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
                Set oDS_PH_PY413A = Nothing
                
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
'                    Call MDC_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY413A", "Code")
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
'                Call PH_PY413_FormItemEnabled
        End Select
    ElseIf (pval.BeforeAction = False) Then
        Select Case pval.MenuUID
            Case "1283":
                oForm.Mode = fm_ADD_MODE
                Call PH_PY413_FormItemEnabled
'                Call PH_PY413_AddMatrixRow
            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//문서찾기
                Call PH_PY413_FormItemEnabled
'                Call PH_PY413_AddMatrixRow
                oForm.Items("Code").CLICK ct_Regular
            Case "1282":    '//문서추가
                Call PH_PY413_FormItemEnabled
'                Call PH_PY413_AddMatrixRow
            Case "1288", "1289", "1290", "1291":
                Call PH_PY413_FormItemEnabled
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


Sub PH_PY413_FormClear()
On Error GoTo PH_PY413_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY413'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PH_PY413_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY413_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY413_DataValidCheck() As Boolean
On Error GoTo PH_PY413_DataValidCheck_Error
    PH_PY413_DataValidCheck = False
    Dim i       As Long
    Dim j       As Long
    
    Dim sQry    As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Exit Function
    
    Set oRecordSet = Nothing
    
    
PH_PY413_DataValidCheck_Error:
    Set oRecordSet = Nothing
    PH_PY413_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY413_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY413_MTX01()

    '//DATA조회
    
    Dim i       As Long
    Dim sQry    As String
    Dim iRow    As Long
    Dim oCombo      As SAPbouiCOM.ComboBox
    
    Dim Param01 As String
    Dim Param02 As String
    Dim Param03 As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY413_MTX01_Error

    Call oForm.Freeze(True)
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    Param01 = Trim(oForm.Items("CLTCOD").Specific.VALUE)
    Param02 = oForm.Items("Year").Specific.VALUE
    Param03 = oForm.Items("MSTCOD").Specific.VALUE
        
    If Trim(Param01) = "" Then
        Call MDC_Com.MDC_GF_Message("사업장이 없습니다. 확인바랍니다..", "E")
        GoTo PH_PY413_MTX01_Exit
    End If
        
    If Trim(Param02) = "" Then
        Call MDC_Com.MDC_GF_Message("년도가 없습니다. 확인바랍니다..", "E")
        GoTo PH_PY413_MTX01_Exit
    End If
    
    If Trim(Param03) = "" Then
        Call MDC_Com.MDC_GF_Message("사번이 없습니다. 확인바랍니다..", "E")
        GoTo PH_PY413_MTX01_Exit
    End If
        
        
        
    sQry = " Select * From [p_seoyhouse] Where saup = '" & Param01 & "' And yyyy = '" & Param02 & "' And sabun = '" & Param03 & "'"
    oRecordSet.DoQuery sQry
    
     'oForm.Items("ws_name1").Specific.VALUE = oRecordSet.Fields("ws_name1").VALUE
     
    oForm.DataSources.UserDataSources.Item("ws_name1").VALUE = oRecordSet.Fields("ws_name1").VALUE
    oForm.DataSources.UserDataSources.Item("ws_name2").VALUE = oRecordSet.Fields("ws_name2").VALUE
    oForm.DataSources.UserDataSources.Item("ws_name3").VALUE = oRecordSet.Fields("ws_name3").VALUE
    oForm.DataSources.UserDataSources.Item("ws_jumin1").VALUE = oRecordSet.Fields("ws_jumin1").VALUE
    oForm.DataSources.UserDataSources.Item("ws_jumin2").VALUE = oRecordSet.Fields("ws_jumin2").VALUE
    oForm.DataSources.UserDataSources.Item("ws_jumin3").VALUE = oRecordSet.Fields("ws_jumin3").VALUE
    Set oCombo = oForm.Items("ws_hcode1").Specific
    oCombo.Select oRecordSet.Fields("ws_hcode1").VALUE, psk_ByValue
    Set oCombo = oForm.Items("ws_hcode2").Specific
    oCombo.Select oRecordSet.Fields("ws_hcode2").VALUE, psk_ByValue
    Set oCombo = oForm.Items("ws_hcode3").Specific
    oCombo.Select oRecordSet.Fields("ws_hcode3").VALUE, psk_ByValue
    oForm.DataSources.UserDataSources.Item("ws_hm1").VALUE = oRecordSet.Fields("ws_hm1").VALUE
    oForm.DataSources.UserDataSources.Item("ws_hm2").VALUE = oRecordSet.Fields("ws_hm2").VALUE
    oForm.DataSources.UserDataSources.Item("ws_hm3").VALUE = oRecordSet.Fields("ws_hm3").VALUE
    oForm.DataSources.UserDataSources.Item("ws_addr1").VALUE = oRecordSet.Fields("ws_addr1").VALUE
    oForm.DataSources.UserDataSources.Item("ws_addr2").VALUE = oRecordSet.Fields("ws_addr2").VALUE
    oForm.DataSources.UserDataSources.Item("ws_addr3").VALUE = oRecordSet.Fields("ws_addr3").VALUE
    oForm.DataSources.UserDataSources.Item("ws_fymd1").VALUE = oRecordSet.Fields("ws_fymd1").VALUE
    oForm.DataSources.UserDataSources.Item("ws_fymd2").VALUE = oRecordSet.Fields("ws_fymd2").VALUE
    oForm.DataSources.UserDataSources.Item("ws_fymd3").VALUE = oRecordSet.Fields("ws_fymd3").VALUE
    oForm.DataSources.UserDataSources.Item("ws_tymd1").VALUE = oRecordSet.Fields("ws_tymd1").VALUE
    oForm.DataSources.UserDataSources.Item("ws_tymd2").VALUE = oRecordSet.Fields("ws_tymd2").VALUE
    oForm.DataSources.UserDataSources.Item("ws_tymd3").VALUE = oRecordSet.Fields("ws_tymd3").VALUE
    oForm.DataSources.UserDataSources.Item("ws_mamt1").VALUE = oRecordSet.Fields("ws_mamt1").VALUE
    oForm.DataSources.UserDataSources.Item("ws_mamt2").VALUE = oRecordSet.Fields("ws_mamt2").VALUE
    oForm.DataSources.UserDataSources.Item("ws_mamt3").VALUE = oRecordSet.Fields("ws_mamt3").VALUE
    oForm.DataSources.UserDataSources.Item("ws_gamt1").VALUE = oRecordSet.Fields("ws_gamt1").VALUE
    oForm.DataSources.UserDataSources.Item("ws_gamt2").VALUE = oRecordSet.Fields("ws_gamt2").VALUE
    oForm.DataSources.UserDataSources.Item("ws_gamt3").VALUE = oRecordSet.Fields("ws_gamt3").VALUE
    
    oForm.DataSources.UserDataSources.Item("dj_name1").VALUE = oRecordSet.Fields("dj_name1").VALUE
    oForm.DataSources.UserDataSources.Item("dj_name2").VALUE = oRecordSet.Fields("dj_name2").VALUE
    oForm.DataSources.UserDataSources.Item("dj_name3").VALUE = oRecordSet.Fields("dj_name3").VALUE
    oForm.DataSources.UserDataSources.Item("dj_jumin1").VALUE = oRecordSet.Fields("dj_jumin1").VALUE
    oForm.DataSources.UserDataSources.Item("dj_jumin2").VALUE = oRecordSet.Fields("dj_jumin2").VALUE
    oForm.DataSources.UserDataSources.Item("dj_jumin3").VALUE = oRecordSet.Fields("dj_jumin3").VALUE
    oForm.DataSources.UserDataSources.Item("dj_fymd1").VALUE = oRecordSet.Fields("dj_fymd1").VALUE
    oForm.DataSources.UserDataSources.Item("dj_fymd2").VALUE = oRecordSet.Fields("dj_fymd2").VALUE
    oForm.DataSources.UserDataSources.Item("dj_fymd3").VALUE = oRecordSet.Fields("dj_fymd3").VALUE
    oForm.DataSources.UserDataSources.Item("dj_tymd1").VALUE = oRecordSet.Fields("dj_tymd1").VALUE
    oForm.DataSources.UserDataSources.Item("dj_tymd2").VALUE = oRecordSet.Fields("dj_tymd2").VALUE
    oForm.DataSources.UserDataSources.Item("dj_tymd3").VALUE = oRecordSet.Fields("dj_tymd3").VALUE
    oForm.DataSources.UserDataSources.Item("dj_eja1").VALUE = oRecordSet.Fields("dj_eja1").VALUE
    oForm.DataSources.UserDataSources.Item("dj_eja2").VALUE = oRecordSet.Fields("dj_eja2").VALUE
    oForm.DataSources.UserDataSources.Item("dj_eja3").VALUE = oRecordSet.Fields("dj_eja3").VALUE
    oForm.DataSources.UserDataSources.Item("dj_tamt1").VALUE = oRecordSet.Fields("dj_tamt1").VALUE
    oForm.DataSources.UserDataSources.Item("dj_tamt2").VALUE = oRecordSet.Fields("dj_tamt2").VALUE
    oForm.DataSources.UserDataSources.Item("dj_tamt3").VALUE = oRecordSet.Fields("dj_tamt3").VALUE
    oForm.DataSources.UserDataSources.Item("dj_wamt1").VALUE = oRecordSet.Fields("dj_wamt1").VALUE
    oForm.DataSources.UserDataSources.Item("dj_wamt2").VALUE = oRecordSet.Fields("dj_wamt2").VALUE
    oForm.DataSources.UserDataSources.Item("dj_wamt3").VALUE = oRecordSet.Fields("dj_wamt3").VALUE
    oForm.DataSources.UserDataSources.Item("dj_eamt1").VALUE = oRecordSet.Fields("dj_eamt1").VALUE
    oForm.DataSources.UserDataSources.Item("dj_eamt2").VALUE = oRecordSet.Fields("dj_eamt2").VALUE
    oForm.DataSources.UserDataSources.Item("dj_eamt3").VALUE = oRecordSet.Fields("dj_eamt3").VALUE
    oForm.DataSources.UserDataSources.Item("dj_gamt1").VALUE = oRecordSet.Fields("dj_gamt1").VALUE
    oForm.DataSources.UserDataSources.Item("dj_gamt2").VALUE = oRecordSet.Fields("dj_gamt2").VALUE
    oForm.DataSources.UserDataSources.Item("dj_gamt3").VALUE = oRecordSet.Fields("dj_gamt3").VALUE
    
    oForm.DataSources.UserDataSources.Item("ld_name1").VALUE = oRecordSet.Fields("ld_name1").VALUE
    oForm.DataSources.UserDataSources.Item("ld_name2").VALUE = oRecordSet.Fields("ld_name2").VALUE
    oForm.DataSources.UserDataSources.Item("ld_name3").VALUE = oRecordSet.Fields("ld_name3").VALUE
    oForm.DataSources.UserDataSources.Item("ld_jumin1").VALUE = oRecordSet.Fields("ld_jumin1").VALUE
    oForm.DataSources.UserDataSources.Item("ld_jumin2").VALUE = oRecordSet.Fields("ld_jumin2").VALUE
    oForm.DataSources.UserDataSources.Item("ld_jumin3").VALUE = oRecordSet.Fields("ld_jumin3").VALUE
    Set oCombo = oForm.Items("ld_hcode1").Specific
    oCombo.Select oRecordSet.Fields("ld_hcode1").VALUE, psk_ByValue
    Set oCombo = oForm.Items("ld_hcode2").Specific
    oCombo.Select oRecordSet.Fields("ld_hcode2").VALUE, psk_ByValue
    Set oCombo = oForm.Items("ld_hcode3").Specific
    oCombo.Select oRecordSet.Fields("ld_hcode3").VALUE, psk_ByValue
    oForm.DataSources.UserDataSources.Item("ld_hm1").VALUE = oRecordSet.Fields("ld_hm1").VALUE
    oForm.DataSources.UserDataSources.Item("ld_hm2").VALUE = oRecordSet.Fields("ld_hm2").VALUE
    oForm.DataSources.UserDataSources.Item("ld_hm3").VALUE = oRecordSet.Fields("ld_hm3").VALUE
    oForm.DataSources.UserDataSources.Item("ld_addr1").VALUE = oRecordSet.Fields("ld_addr1").VALUE
    oForm.DataSources.UserDataSources.Item("ld_addr2").VALUE = oRecordSet.Fields("ld_addr2").VALUE
    oForm.DataSources.UserDataSources.Item("ld_addr3").VALUE = oRecordSet.Fields("ld_addr3").VALUE
    oForm.DataSources.UserDataSources.Item("ld_fymd1").VALUE = oRecordSet.Fields("ld_fymd1").VALUE
    oForm.DataSources.UserDataSources.Item("ld_fymd2").VALUE = oRecordSet.Fields("ld_fymd2").VALUE
    oForm.DataSources.UserDataSources.Item("ld_fymd3").VALUE = oRecordSet.Fields("ld_fymd3").VALUE
    oForm.DataSources.UserDataSources.Item("ld_tymd1").VALUE = oRecordSet.Fields("ld_tymd1").VALUE
    oForm.DataSources.UserDataSources.Item("ld_tymd2").VALUE = oRecordSet.Fields("ld_tymd2").VALUE
    oForm.DataSources.UserDataSources.Item("ld_tymd3").VALUE = oRecordSet.Fields("ld_tymd3").VALUE
    oForm.DataSources.UserDataSources.Item("ld_bamt1").VALUE = oRecordSet.Fields("ld_bamt1").VALUE
    oForm.DataSources.UserDataSources.Item("ld_bamt2").VALUE = oRecordSet.Fields("ld_bamt2").VALUE
    oForm.DataSources.UserDataSources.Item("ld_bamt3").VALUE = oRecordSet.Fields("ld_bamt3").VALUE
    
    
    oForm.ActiveItem = "ws_name1"
    
    oForm.Update
    
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Sub
PH_PY413_MTX01_Exit:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Sub
PH_PY413_MTX01_Error:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY413_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub
Private Sub PH_PY413_MTX02(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)


'    '//그리드 자료를 head에 로드
'
'    Dim i       As Long
'    Dim sQry    As String
'    Dim sRow As Long
'
'    Dim Param01 As String
'    Dim Param02 As String
'    Dim Param03 As String
'    Dim Param04 As String
'    Dim Param05 As String
'    Dim Param06 As String
'    Dim Param07 As String
'
'    Dim oCombo      As SAPbouiCOM.ComboBox
'    Dim oRecordSet As SAPbobsCOM.Recordset
'
'    On Error GoTo PH_PY413_MTX02_Error
'
'    Call oForm.Freeze(True)
'    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
'
'    sRow = oRow
'
'
'    Param01 = Trim(oForm.Items("CLTCOD").Specific.VALUE)
'    Param02 = oDS_PH_PY413A.Columns.Item("연도").Cells(oRow).VALUE
'    Param03 = oDS_PH_PY413A.Columns.Item("사번").Cells(oRow).VALUE
'    Param04 = oDS_PH_PY413A.Columns.Item("주민번호").Cells(oRow).VALUE
'    Param05 = oDS_PH_PY413A.Columns.Item("지급처상호").Cells(oRow).VALUE
'    Param06 = oDS_PH_PY413A.Columns.Item("지급일자").Cells(oRow).VALUE
'    Param07 = oDS_PH_PY413A.Columns.Item("사업자번호").Cells(oRow).VALUE
'
'
'
'    sQry = "EXEC PH_PY413_02 '" & Param01 & "', '" & Param02 & "', '" & Param03 & "', '" & Param04 & "', '" & Param05 & "', '" & Param06 & "', '" & Param07 & "'"
'    Call oRecordSet.DoQuery(sQry)
'
'    If (oRecordSet.RecordCount = 0) Then
'
'        'oForm.Items("MSTCOD").Specific.VALUE = oDS_PH_PY413A.Columns.Item("MSTCOD").Cells(oRow).VALUE
'        'oForm.Items("FullName").Specific.VALUE = oDS_PH_PY413A.Columns.Item("FullName").Cells(oRow).VALUE
'
'        oForm.Items("kname").Specific.VALUE = ""
'        oForm.Items("juminno").Specific.VALUE = ""
'        oForm.Items("custnm").Specific.VALUE = ""
'        oForm.Items("entno").Specific.VALUE = ""
'        oForm.Items("payymd").Specific.VALUE = ""
'
'        oForm.Items("medcex").Specific.VALUE = 0
'        oForm.Items("ntamt").Specific.VALUE = 0
'        oForm.Items("cont").Specific.VALUE = 0
'
'        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
'        GoTo PH_PY413_MTX02_Exit
'    End If
'
'    Set oCombo = oForm.Items("rel").Specific
'    oCombo.Select oRecordSet.Fields("rel").VALUE, psk_ByValue
'
'    oForm.DataSources.UserDataSources.Item("kname").VALUE = oRecordSet.Fields("kname").VALUE
'    oForm.DataSources.UserDataSources.Item("juminno").VALUE = oRecordSet.Fields("juminno").VALUE
'
'    Set oCombo = oForm.Items("empdiv").Specific
'    oCombo.Select oRecordSet.Fields("empdiv").VALUE, psk_ByValue
'
'    oForm.DataSources.UserDataSources.Item("custnm").VALUE = oRecordSet.Fields("custnm").VALUE
'    oForm.DataSources.UserDataSources.Item("entno").VALUE = oRecordSet.Fields("entno").VALUE
'    oForm.DataSources.UserDataSources.Item("payymd").VALUE = oRecordSet.Fields("payymd").VALUE
'
'    Set oCombo = oForm.Items("gubun").Specific
'    oCombo.Select oRecordSet.Fields("gubun").VALUE, psk_ByValue
'
'    oForm.DataSources.UserDataSources.Item("medcex").VALUE = oRecordSet.Fields("medcex").VALUE
'    oForm.DataSources.UserDataSources.Item("ntamt").VALUE = oRecordSet.Fields("ntamt").VALUE
'    oForm.DataSources.UserDataSources.Item("cont").VALUE = oRecordSet.Fields("cont").VALUE
'
'    Set oCombo = oForm.Items("olddiv").Specific
'    oCombo.Select oRecordSet.Fields("olddiv").VALUE, psk_ByValue
'
'    Set oCombo = oForm.Items("deform").Specific
'    oCombo.Select oRecordSet.Fields("deform").VALUE, psk_ByValue
'
''    '//부서
''    oForm.Items("TeamName").Specific.VALUE = oRecordSet.Fields("TeamName").VALUE
''    oForm.Items("RspName").Specific.VALUE = oRecordSet.Fields("RspName").VALUE
''    oForm.Items("ClsName").Specific.VALUE = oRecordSet.Fields("ClsName").VALUE
'
'    '//Key Disable
'    oForm.Items("CLTCOD").Enabled = False
'    oForm.Items("Year").Enabled = False
'    oForm.Items("MSTCOD").Enabled = False
'
'    oForm.Items("juminno").Enabled = False
'    oForm.Items("custnm").Enabled = False
'    oForm.Items("payymd").Enabled = False
'    oForm.Items("entno").Enabled = False
'
'
'    oForm.Update
'
'    Set oRecordSet = Nothing
'    Call oForm.Freeze(False)
'    Exit Sub
'PH_PY413_MTX02_Exit:
'    Set oRecordSet = Nothing
'    Call oForm.Freeze(False)
'
'    Exit Sub
'PH_PY413_MTX02_Error:
'    Set oRecordSet = Nothing
'    Call oForm.Freeze(False)
'    Sbo_Application.SetStatusBarMessage "PH_PY413_MTX02_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY413_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY413_Validate_Error
    PH_PY413_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY413A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", bmt_Short, True
        PH_PY413_Validate = False
        GoTo PH_PY413_Validate_Exit
    End If
'
    If ValidateType = "수정" Then

    ElseIf ValidateType = "행삭제" Then

    ElseIf ValidateType = "취소" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY413_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY413_Validate_Error:
    PH_PY413_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY413_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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


Private Sub PH_PY413_SAVE()

    '//데이타 저장
    
    Dim i       As Long
    Dim sQry    As String
    
    Dim saup$, yyyy$, sabun$
    Dim ws_name1$, ws_name2$, ws_name3$, ws_jumin1$, ws_jumin2$, ws_jumin3$, ws_addr1$, ws_addr2$, ws_addr3$
    Dim ws_fymd1$, ws_fymd2$, ws_fymd3$, ws_tymd1$, ws_tymd2$, ws_tymd3$
    Dim ws_mamt1, ws_mamt2, ws_mamt3, ws_gamt1, ws_gamt2, ws_gamt3  As Double
    Dim ws_hcode1$, ws_hcode2$, ws_hcode3$
    Dim ws_hm1, ws_hm2, ws_hm3 As Double
    
    Dim dj_name1$, dj_name2$, dj_name3$, dj_jumin1$, dj_jumin2$, dj_jumin3$
    Dim dj_fymd1$, dj_fymd2$, dj_fymd3$, dj_tymd1$, dj_tymd2$, dj_tymd3$
    Dim dj_eja1$, dj_eja2$, dj_eja3$
    Dim dj_tamt1, dj_tamt2, dj_tamt3, dj_wamt1, dj_wamt2, dj_wamt3  As Double
    Dim dj_eamt1, dj_eamt2, dj_eamt3, dj_gamt1, dj_gamt2, dj_gamt3  As Double
    
    Dim ld_name1$, ld_name2$, ld_name3$, ld_jumin1$, ld_jumin2$, ld_jumin3$, ld_addr1$, ld_addr2$, ld_addr3$
    Dim ld_fymd1$, ld_fymd2$, ld_fymd3$, ld_tymd1$, ld_tymd2$, ld_tymd3$
    Dim ld_bamt1, ld_bamt2, ld_bamt3 As Double
    Dim ld_hcode1$, ld_hcode2$, ld_hcode3$
    Dim ld_hm1, ld_hm2, ld_hm3 As Double
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    saup = Trim(oForm.Items("CLTCOD").Specific.VALUE)
    yyyy = oForm.Items("Year").Specific.VALUE
    sabun = Trim(Str(oForm.Items("MSTCOD").Specific.VALUE))
        
    ws_name1 = oForm.Items("ws_name1").Specific.VALUE
    ws_name2 = oForm.Items("ws_name2").Specific.VALUE
    ws_name3 = oForm.Items("ws_name3").Specific.VALUE
    
    ws_jumin1 = oForm.Items("ws_jumin1").Specific.VALUE
    ws_jumin2 = oForm.Items("ws_jumin2").Specific.VALUE
    ws_jumin3 = oForm.Items("ws_jumin3").Specific.VALUE
    
    ws_hcode1 = oForm.Items("ws_hcode1").Specific.VALUE
    ws_hcode2 = oForm.Items("ws_hcode2").Specific.VALUE
    ws_hcode3 = oForm.Items("ws_hcode3").Specific.VALUE
    
    ws_hm1 = oForm.Items("ws_hm1").Specific.VALUE
    ws_hm2 = oForm.Items("ws_hm2").Specific.VALUE
    ws_hm3 = oForm.Items("ws_hm3").Specific.VALUE
    
    ws_addr1 = oForm.Items("ws_addr1").Specific.VALUE
    ws_addr2 = oForm.Items("ws_addr2").Specific.VALUE
    ws_addr3 = oForm.Items("ws_addr3").Specific.VALUE
    
    ws_fymd1 = oForm.Items("ws_fymd1").Specific.VALUE
    ws_fymd2 = oForm.Items("ws_fymd2").Specific.VALUE
    ws_fymd3 = oForm.Items("ws_fymd3").Specific.VALUE
    
    ws_tymd1 = oForm.Items("ws_tymd1").Specific.VALUE
    ws_tymd2 = oForm.Items("ws_tymd2").Specific.VALUE
    ws_tymd3 = oForm.Items("ws_tymd3").Specific.VALUE
    
    ws_mamt1 = oForm.Items("ws_mamt1").Specific.VALUE
    ws_mamt2 = oForm.Items("ws_mamt2").Specific.VALUE
    ws_mamt3 = oForm.Items("ws_mamt3").Specific.VALUE
    
    ws_gamt1 = oForm.Items("ws_gamt1").Specific.VALUE
    ws_gamt2 = oForm.Items("ws_gamt2").Specific.VALUE
    ws_gamt3 = oForm.Items("ws_gamt3").Specific.VALUE
    
    dj_name1 = oForm.Items("dj_name1").Specific.VALUE
    dj_name2 = oForm.Items("dj_name2").Specific.VALUE
    dj_name3 = oForm.Items("dj_name3").Specific.VALUE
    
    dj_jumin1 = oForm.Items("dj_jumin1").Specific.VALUE
    dj_jumin2 = oForm.Items("dj_jumin2").Specific.VALUE
    dj_jumin3 = oForm.Items("dj_jumin3").Specific.VALUE
    
    dj_fymd1 = oForm.Items("dj_fymd1").Specific.VALUE
    dj_fymd2 = oForm.Items("dj_fymd2").Specific.VALUE
    dj_fymd3 = oForm.Items("dj_fymd3").Specific.VALUE
    
    dj_tymd1 = oForm.Items("dj_tymd1").Specific.VALUE
    dj_tymd2 = oForm.Items("dj_tymd2").Specific.VALUE
    dj_tymd3 = oForm.Items("dj_tymd3").Specific.VALUE
    
    dj_eja1 = oForm.Items("dj_eja1").Specific.VALUE
    dj_eja2 = oForm.Items("dj_eja2").Specific.VALUE
    dj_eja3 = oForm.Items("dj_eja3").Specific.VALUE
    
    dj_tamt1 = oForm.Items("dj_tamt1").Specific.VALUE
    dj_tamt2 = oForm.Items("dj_tamt2").Specific.VALUE
    dj_tamt3 = oForm.Items("dj_tamt3").Specific.VALUE
    
    dj_wamt1 = oForm.Items("dj_wamt1").Specific.VALUE
    dj_wamt2 = oForm.Items("dj_wamt2").Specific.VALUE
    dj_wamt3 = oForm.Items("dj_wamt3").Specific.VALUE
    
    dj_eamt1 = oForm.Items("dj_eamt1").Specific.VALUE
    dj_eamt2 = oForm.Items("dj_eamt2").Specific.VALUE
    dj_eamt3 = oForm.Items("dj_eamt3").Specific.VALUE
    
    dj_gamt1 = oForm.Items("dj_gamt1").Specific.VALUE
    dj_gamt2 = oForm.Items("dj_gamt2").Specific.VALUE
    dj_gamt3 = oForm.Items("dj_gamt3").Specific.VALUE
    
    ld_name1 = oForm.Items("ld_name1").Specific.VALUE
    ld_name2 = oForm.Items("ld_name2").Specific.VALUE
    ld_name3 = oForm.Items("ld_name3").Specific.VALUE
    
    ld_jumin1 = oForm.Items("ld_jumin1").Specific.VALUE
    ld_jumin2 = oForm.Items("ld_jumin2").Specific.VALUE
    ld_jumin3 = oForm.Items("ld_jumin3").Specific.VALUE
    
    ld_hcode1 = oForm.Items("ld_hcode1").Specific.VALUE
    ld_hcode2 = oForm.Items("ld_hcode2").Specific.VALUE
    ld_hcode3 = oForm.Items("ld_hcode3").Specific.VALUE
    
    ld_hm1 = oForm.Items("ld_hm1").Specific.VALUE
    ld_hm2 = oForm.Items("ld_hm2").Specific.VALUE
    ld_hm3 = oForm.Items("ld_hm3").Specific.VALUE
    
    ld_addr1 = oForm.Items("ld_addr1").Specific.VALUE
    ld_addr2 = oForm.Items("ld_addr2").Specific.VALUE
    ld_addr3 = oForm.Items("ld_addr3").Specific.VALUE
    
    ld_fymd1 = oForm.Items("ld_fymd1").Specific.VALUE
    ld_fymd2 = oForm.Items("ld_fymd2").Specific.VALUE
    ld_fymd3 = oForm.Items("ld_fymd3").Specific.VALUE
    
    ld_tymd1 = oForm.Items("ld_tymd1").Specific.VALUE
    ld_tymd2 = oForm.Items("ld_tymd2").Specific.VALUE
    ld_tymd3 = oForm.Items("ld_tymd3").Specific.VALUE
    
    ld_bamt1 = oForm.Items("ld_bamt1").Specific.VALUE
    ld_bamt2 = oForm.Items("ld_bamt2").Specific.VALUE
    ld_bamt3 = oForm.Items("ld_bamt3").Specific.VALUE
    
    On Error GoTo PH_PY413_SAVE_Error

    Call oForm.Freeze(True)
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If Trim(yyyy) = "" Then
        Call MDC_Com.MDC_GF_Message("년도가 없습니다. 확인바랍니다..", "E")
        GoTo PH_PY413_SAVE_Exit
    End If
    
    If Trim(saup) = "" Then
        Call MDC_Com.MDC_GF_Message("사업장이 없습니다. 확인바랍니다..", "E")
        GoTo PH_PY413_SAVE_Exit
    End If
    
    If Trim(sabun) = "" Then
        Call MDC_Com.MDC_GF_Message("사번이 없습니다. 확인바랍니다..", "E")
        GoTo PH_PY413_SAVE_Exit
    End If
            
    sQry = " Select Count(*) From [p_seoyhouse] Where saup = '" & saup & "' And yyyy = '" & yyyy & "' And sabun = '" & sabun & "'"
    oRecordSet.DoQuery sQry
    
    If oRecordSet.Fields(0).VALUE > 0 Then
        '//갱신
        
        sQry = "Update [p_seoyhouse] set "
        sQry = sQry + "ws_name1 = '" & ws_name1 & "',"
        sQry = sQry + "ws_name2 = '" & ws_name2 & "',"
        sQry = sQry + "ws_name3 = '" & ws_name3 & "',"
        sQry = sQry + "ws_jumin1 = '" & ws_jumin1 & "',"
        sQry = sQry + "ws_jumin2 = '" & ws_jumin2 & "',"
        sQry = sQry + "ws_jumin3 = '" & ws_jumin3 & "',"
        sQry = sQry + "ws_hcode1 = '" & ws_hcode1 & "',"
        sQry = sQry + "ws_hcode2 = '" & ws_hcode2 & "',"
        sQry = sQry + "ws_hcode3 = '" & ws_hcode3 & "',"
        sQry = sQry + "ws_hm1 = '" & ws_hm1 & "',"
        sQry = sQry + "ws_hm2 = '" & ws_hm2 & "',"
        sQry = sQry + "ws_hm3 = '" & ws_hm3 & "',"
        sQry = sQry + "ws_addr1 = '" & ws_addr1 & "',"
        sQry = sQry + "ws_addr2 = '" & ws_addr2 & "',"
        sQry = sQry + "ws_addr3 = '" & ws_addr3 & "',"
        sQry = sQry + "ws_fymd1 = '" & ws_fymd1 & "',"
        sQry = sQry + "ws_fymd2 = '" & ws_fymd2 & "',"
        sQry = sQry + "ws_fymd3 = '" & ws_fymd3 & "',"
        sQry = sQry + "ws_tymd1 = '" & ws_tymd1 & "',"
        sQry = sQry + "ws_tymd2 = '" & ws_tymd2 & "',"
        sQry = sQry + "ws_tymd3 = '" & ws_tymd3 & "',"
        sQry = sQry + "ws_mamt1 = " & ws_mamt1 & ","
        sQry = sQry + "ws_mamt2 = " & ws_mamt2 & ","
        sQry = sQry + "ws_mamt3 = " & ws_mamt3 & ","
        sQry = sQry + "ws_gamt1 = " & ws_gamt1 & ","
        sQry = sQry + "ws_gamt2 = " & ws_gamt2 & ","
        sQry = sQry + "ws_gamt3 = " & ws_gamt3 & ","
        
        sQry = sQry + "dj_name1 = '" & dj_name1 & "',"
        sQry = sQry + "dj_name2 = '" & dj_name2 & "',"
        sQry = sQry + "dj_name3 = '" & dj_name3 & "',"
        sQry = sQry + "dj_jumin1 = '" & dj_jumin1 & "',"
        sQry = sQry + "dj_jumin2 = '" & dj_jumin2 & "',"
        sQry = sQry + "dj_jumin3 = '" & dj_jumin3 & "',"
        sQry = sQry + "dj_fymd1 = '" & dj_fymd1 & "',"
        sQry = sQry + "dj_fymd2 = '" & dj_fymd2 & "',"
        sQry = sQry + "dj_fymd3 = '" & dj_fymd3 & "',"
        sQry = sQry + "dj_tymd1 = '" & dj_tymd1 & "',"
        sQry = sQry + "dj_tymd2 = '" & dj_tymd2 & "',"
        sQry = sQry + "dj_tymd3 = '" & dj_tymd3 & "',"
        sQry = sQry + "dj_eja1 = '" & dj_eja1 & "',"
        sQry = sQry + "dj_eja2 = '" & dj_eja2 & "',"
        sQry = sQry + "dj_eja3 = '" & dj_eja3 & "',"
        sQry = sQry + "dj_tamt1 = " & dj_tamt1 & ","
        sQry = sQry + "dj_tamt2 = " & dj_tamt2 & ","
        sQry = sQry + "dj_tamt3 = " & dj_tamt3 & ","
        sQry = sQry + "dj_wamt1 = " & dj_wamt1 & ","
        sQry = sQry + "dj_wamt2 = " & dj_wamt2 & ","
        sQry = sQry + "dj_wamt3 = " & dj_wamt3 & ","
        sQry = sQry + "dj_eamt1 = " & dj_eamt1 & ","
        sQry = sQry + "dj_eamt2 = " & dj_eamt2 & ","
        sQry = sQry + "dj_eamt3 = " & dj_eamt3 & ","
        sQry = sQry + "dj_gamt1 = " & dj_gamt1 & ","
        sQry = sQry + "dj_gamt2 = " & dj_gamt2 & ","
        sQry = sQry + "dj_gamt3 = " & dj_gamt3 & ","
        
        sQry = sQry + "ld_name1 = '" & ld_name1 & "',"
        sQry = sQry + "ld_name2 = '" & ld_name2 & "',"
        sQry = sQry + "ld_name3 = '" & ld_name3 & "',"
        sQry = sQry + "ld_jumin1 = '" & ld_jumin1 & "',"
        sQry = sQry + "ld_jumin2 = '" & ld_jumin2 & "',"
        sQry = sQry + "ld_jumin3 = '" & ld_jumin3 & "',"
        sQry = sQry + "ld_hcode1 = '" & ld_hcode1 & "',"
        sQry = sQry + "ld_hcode2 = '" & ld_hcode2 & "',"
        sQry = sQry + "ld_hcode3 = '" & ld_hcode3 & "',"
        sQry = sQry + "ld_hm1 = '" & ld_hm1 & "',"
        sQry = sQry + "ld_hm2 = '" & ld_hm2 & "',"
        sQry = sQry + "ld_hm3 = '" & ld_hm3 & "',"
        sQry = sQry + "ld_addr1 = '" & ld_addr1 & "',"
        sQry = sQry + "ld_addr2 = '" & ld_addr2 & "',"
        sQry = sQry + "ld_addr3 = '" & ld_addr3 & "',"
        sQry = sQry + "ld_fymd1 = '" & ld_fymd1 & "',"
        sQry = sQry + "ld_fymd2 = '" & ld_fymd2 & "',"
        sQry = sQry + "ld_fymd3 = '" & ld_fymd3 & "',"
        sQry = sQry + "ld_tymd1 = '" & ld_tymd1 & "',"
        sQry = sQry + "ld_tymd2 = '" & ld_tymd2 & "',"
        sQry = sQry + "ld_tymd3 = '" & ld_tymd3 & "',"
        sQry = sQry + "ld_bamt1 = " & ld_bamt1 & ","
        sQry = sQry + "ld_bamt2 = " & ld_bamt2 & ","
        sQry = sQry + "ld_bamt3 = " & ld_bamt3 & ""
        
        sQry = sQry + " Where saup = '" & saup & "' And yyyy = '" & yyyy & "' And sabun = '" & sabun & "'"
        
        oRecordSet.DoQuery sQry
        
    Else
        
        '//신규
        sQry = "INSERT INTO [p_seoyhouse]"
        sQry = sQry & " ("
        sQry = sQry & "saup,"
        sQry = sQry & "yyyy,"
        sQry = sQry & "sabun,"
        sQry = sQry & "ws_name1, "
        sQry = sQry & "ws_name2, "
        sQry = sQry & "ws_name3, "
        sQry = sQry & "ws_jumin1, "
        sQry = sQry & "ws_jumin2, "
        sQry = sQry & "ws_jumin3, "
        sQry = sQry & "ws_hcode1, "
        sQry = sQry & "ws_hcode2, "
        sQry = sQry & "ws_hcode3, "
        sQry = sQry & "ws_hm1, "
        sQry = sQry & "ws_hm2, "
        sQry = sQry & "ws_hm3, "
        sQry = sQry & "ws_addr1, "
        sQry = sQry & "ws_addr2, "
        sQry = sQry & "ws_addr3, "
        sQry = sQry & "ws_fymd1, "
        sQry = sQry & "ws_fymd2, "
        sQry = sQry & "ws_fymd3, "
        sQry = sQry & "ws_tymd1, "
        sQry = sQry & "ws_tymd2, "
        sQry = sQry & "ws_tymd3, "
        sQry = sQry & "ws_mamt1, "
        sQry = sQry & "ws_mamt2, "
        sQry = sQry & "ws_mamt3, "
        sQry = sQry & "ws_gamt1, "
        sQry = sQry & "ws_gamt2, "
        sQry = sQry & "ws_gamt3, "
        
        sQry = sQry & "dj_name1, "
        sQry = sQry & "dj_name2, "
        sQry = sQry & "dj_name3, "
        sQry = sQry & "dj_jumin1, "
        sQry = sQry & "dj_jumin2, "
        sQry = sQry & "dj_jumin3, "
        sQry = sQry & "dj_fymd1, "
        sQry = sQry & "dj_fymd2, "
        sQry = sQry & "dj_fymd3, "
        sQry = sQry & "dj_tymd1, "
        sQry = sQry & "dj_tymd2, "
        sQry = sQry & "dj_tymd3, "
        sQry = sQry & "dj_eja1, "
        sQry = sQry & "dj_eja2, "
        sQry = sQry & "dj_eja3, "
        sQry = sQry & "dj_tamt1, "
        sQry = sQry & "dj_tamt2, "
        sQry = sQry & "dj_tamt3, "
        sQry = sQry & "dj_wamt1, "
        sQry = sQry & "dj_wamt2, "
        sQry = sQry & "dj_wamt3, "
        sQry = sQry & "dj_eamt1, "
        sQry = sQry & "dj_eamt2, "
        sQry = sQry & "dj_eamt3, "
        sQry = sQry & "dj_gamt1, "
        sQry = sQry & "dj_gamt2, "
        sQry = sQry & "dj_gamt3, "
        
        sQry = sQry & "ld_name1, "
        sQry = sQry & "ld_name2, "
        sQry = sQry & "ld_name3, "
        sQry = sQry & "ld_jumin1, "
        sQry = sQry & "ld_jumin2, "
        sQry = sQry & "ld_jumin3, "
        sQry = sQry & "ld_hcode1, "
        sQry = sQry & "ld_hcode2, "
        sQry = sQry & "ld_hcode3, "
        sQry = sQry & "ld_hm1, "
        sQry = sQry & "ld_hm2, "
        sQry = sQry & "ld_hm3, "
        sQry = sQry & "ld_addr1, "
        sQry = sQry & "ld_addr2, "
        sQry = sQry & "ld_addr3, "
        sQry = sQry & "ld_fymd1, "
        sQry = sQry & "ld_fymd2, "
        sQry = sQry & "ld_fymd3, "
        sQry = sQry & "ld_tymd1, "
        sQry = sQry & "ld_tymd2, "
        sQry = sQry & "ld_tymd3, "
        sQry = sQry & "ld_bamt1, "
        sQry = sQry & "ld_bamt2, "
        sQry = sQry & "ld_bamt3 "
        sQry = sQry & " ) "
        sQry = sQry & "VALUES("

        sQry = sQry & "'" & saup & "',"
        sQry = sQry & "'" & yyyy & "',"
        sQry = sQry & "'" & sabun & "',"
        sQry = sQry & "'" & ws_name1 & "',"
        sQry = sQry & "'" & ws_name2 & "',"
        sQry = sQry & "'" & ws_name3 & "',"
        sQry = sQry & "'" & ws_jumin1 & "',"
        sQry = sQry & "'" & ws_jumin2 & "',"
        sQry = sQry & "'" & ws_jumin3 & "',"
        sQry = sQry & "'" & ws_hcode1 & "',"
        sQry = sQry & "'" & ws_hcode2 & "',"
        sQry = sQry & "'" & ws_hcode3 & "',"
        sQry = sQry & ws_hm1 & ","
        sQry = sQry & ws_hm2 & ","
        sQry = sQry & ws_hm3 & ","
        sQry = sQry & "'" & ws_addr1 & "',"
        sQry = sQry & "'" & ws_addr2 & "',"
        sQry = sQry & "'" & ws_addr3 & "',"
        sQry = sQry & "'" & ws_fymd1 & "',"
        sQry = sQry & "'" & ws_fymd2 & "',"
        sQry = sQry & "'" & ws_fymd3 & "',"
        sQry = sQry & "'" & ws_tymd1 & "',"
        sQry = sQry & "'" & ws_tymd2 & "',"
        sQry = sQry & "'" & ws_tymd3 & "',"
        sQry = sQry & ws_mamt1 & ","
        sQry = sQry & ws_mamt2 & ","
        sQry = sQry & ws_mamt3 & ","
        sQry = sQry & ws_gamt1 & ","
        sQry = sQry & ws_gamt2 & ","
        sQry = sQry & ws_gamt3 & ","
        
        sQry = sQry & "'" & dj_name1 & "',"
        sQry = sQry & "'" & dj_name2 & "',"
        sQry = sQry & "'" & dj_name3 & "',"
        sQry = sQry & "'" & dj_jumin1 & "',"
        sQry = sQry & "'" & dj_jumin2 & "',"
        sQry = sQry & "'" & dj_jumin3 & "',"
        sQry = sQry & "'" & dj_fymd1 & "',"
        sQry = sQry & "'" & dj_fymd2 & "',"
        sQry = sQry & "'" & dj_fymd3 & "',"
        sQry = sQry & "'" & dj_tymd1 & "',"
        sQry = sQry & "'" & dj_tymd2 & "',"
        sQry = sQry & "'" & dj_tymd3 & "',"
        sQry = sQry & "'" & dj_eja1 & "',"
        sQry = sQry & "'" & dj_eja2 & "',"
        sQry = sQry & "'" & dj_eja3 & "',"
        sQry = sQry & dj_tamt1 & ","
        sQry = sQry & dj_tamt2 & ","
        sQry = sQry & dj_tamt3 & ","
        sQry = sQry & dj_wamt1 & ","
        sQry = sQry & dj_wamt2 & ","
        sQry = sQry & dj_wamt3 & ","
        sQry = sQry & dj_eamt1 & ","
        sQry = sQry & dj_eamt2 & ","
        sQry = sQry & dj_eamt3 & ","
        sQry = sQry & dj_gamt1 & ","
        sQry = sQry & dj_gamt2 & ","
        sQry = sQry & dj_gamt3 & ","
        
        sQry = sQry & "'" & ld_name1 & "',"
        sQry = sQry & "'" & ld_name2 & "',"
        sQry = sQry & "'" & ld_name3 & "',"
        sQry = sQry & "'" & ld_jumin1 & "',"
        sQry = sQry & "'" & ld_jumin2 & "',"
        sQry = sQry & "'" & ld_jumin3 & "',"
        sQry = sQry & "'" & ld_hcode1 & "',"
        sQry = sQry & "'" & ld_hcode2 & "',"
        sQry = sQry & "'" & ld_hcode3 & "',"
        sQry = sQry & ld_hm1 & ","
        sQry = sQry & ld_hm2 & ","
        sQry = sQry & ld_hm3 & ","
        sQry = sQry & "'" & ld_addr1 & "',"
        sQry = sQry & "'" & ld_addr2 & "',"
        sQry = sQry & "'" & ld_addr3 & "',"
        sQry = sQry & "'" & ld_fymd1 & "',"
        sQry = sQry & "'" & ld_fymd2 & "',"
        sQry = sQry & "'" & ld_fymd3 & "',"
        sQry = sQry & "'" & ld_tymd1 & "',"
        sQry = sQry & "'" & ld_tymd2 & "',"
        sQry = sQry & "'" & ld_tymd3 & "',"
        sQry = sQry & ld_bamt1 & ","
        sQry = sQry & ld_bamt2 & ","
        sQry = sQry & ld_bamt3
        sQry = sQry & " ) "

        oRecordSet.DoQuery sQry
    End If
    
    
    Call PH_PY413_FormItemEnabled
    
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)

    Call PH_PY413_MTX01

    Exit Sub
    
PH_PY413_SAVE_Exit:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    
    Exit Sub
PH_PY413_SAVE_Error:
    Call oForm.Freeze(False)
   
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "PH_PY413_SAVE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub PH_PY413_Delete()
    '//선택된 자료 삭제
    
    Dim saup$, yyyy$, sabun$, kname$
    
    Dim i As Integer
    Dim cnt As Integer
    
    Dim sQry    As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY413_Delete_Error

    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Call oForm.Freeze(True)
    
    kname = oForm.Items("FullName").Specific.VALUE
    
    saup = Trim(oForm.Items("CLTCOD").Specific.VALUE)
    yyyy = oForm.Items("Year").Specific.VALUE
    sabun = Trim(Str(oForm.Items("MSTCOD").Specific.VALUE))
        
    sQry = " Select Count(*) From [p_seoyhouse] Where saup = '" & saup & "' And yyyy = '" & yyyy & "' And sabun = '" & sabun & "'"
    oRecordSet.DoQuery sQry
    
    cnt = oRecordSet.Fields(0).VALUE
    If cnt > 0 Then
    
        If Trim(yyyy) = "" Then
            Call MDC_Com.MDC_GF_Message("년도가 없습니다. 확인바랍니다..", "E")
            GoTo PH_PY413_Delete_Exit
        End If
        
        If Trim(saup) = "" Then
            Call MDC_Com.MDC_GF_Message("사업장이 없습니다. 확인바랍니다..", "E")
            GoTo PH_PY413_Delete_Exit
        End If
        If Trim(sabun) = "" Then
            Call MDC_Com.MDC_GF_Message("사번이 없습니다. 확인바랍니다..", "E")
            GoTo PH_PY413_Delete_Exit
        End If
        
        
        If Sbo_Application.MessageBox(" 선택한사원('" & kname & "')의 모든자료를 삭제하시겠습니까? ?", "2", "예", "아니오") = "1" Then
            sQry = "Delete From [p_seoyhouse] Where saup = '" & saup & "' And yyyy = '" & yyyy & "' And sabun = '" & sabun & "'"
            oRecordSet.DoQuery sQry
        End If
    End If
    
    Call oForm.Freeze(False)
        
    Call PH_PY413_MTX01
    
    oForm.ActiveItem = "MSTCOD"
    
    Set oRecordSet = Nothing
    
    
    Exit Sub
PH_PY413_Delete_Exit:
    Set oRecordSet = Nothing
        
    Call oForm.Freeze(False)
    Exit Sub
PH_PY413_Delete_Error:
    Set oRecordSet = Nothing
   
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY413_Delete_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub PH_PY413_TitleSetting(iRow As Long)
'    Dim i               As Long
'    Dim j               As Long
'    Dim sQry            As String
'
'    Dim COLNAM(5)       As String
'
'    Dim oColumn         As SAPbouiCOM.EditTextColumn
'    Dim oComboCol       As SAPbouiCOM.ComboBoxColumn
'
'    Dim oRecordSet  As SAPbobsCOM.Recordset
'
'    On Error GoTo Error_Message
'
'    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
'
'    oForm.Freeze True
'
'    COLNAM(0) = "년도"
'    COLNAM(1) = "부서"
'    COLNAM(2) = "담당"
'    COLNAM(3) = "사번"
'    COLNAM(4) = "성명"
'    COLNAM(5) = "직급"
'
'    For i = 0 To UBound(COLNAM)
'        oGrid1.Columns.Item(i).TitleObject.Caption = COLNAM(i)
'        oGrid1.Columns.Item(i).Editable = False
'
'        oGrid1.Columns.Item(i).RightJustified = True
'
'    Next i
'
'    oGrid1.AutoResizeColumns
'
'    oForm.Freeze False
'
'    Set oColumn = Nothing
'
'    Exit Sub
'
'Error_Message:
'    oForm.Freeze False
'    Set oColumn = Nothing
'    Sbo_Application.SetStatusBarMessage "PH_PY413_TitleSetting Error : " & Space(10) & Err.Description, bmt_Short, True
End Sub









