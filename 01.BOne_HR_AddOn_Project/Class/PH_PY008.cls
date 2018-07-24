VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY008"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY008.cls
'//  Module         : 근태관리 > 근태등록
'//  Desc           : 근태등록
'//********************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form

Public oGrid1           As SAPbouiCOM.Grid
Public oDS_PH_PY008     As SAPbouiCOM.DataTable


Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long
Private oWorkType       As String
Private oRspCodeYN      As String  '담당수정 권한


Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY008.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY008_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY008"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
'    oForm.DataBrowser.BrowseBy = "Code"
        
    oForm.PaneLevel = 1
    oForm.Freeze True
    Call PH_PY008_CreateItems
    Call PH_PY008_FormItemEnabled
    Call PH_PY008_EnableMenus
'    Call PH_PY008_SetDocument(oFromDocEntry01)
'    Call PH_PY008_FormResize
    
    
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

Private Function PH_PY008_CreateItems() As Boolean

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
    
    On Error GoTo PH_PY008_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Set oGrid1 = oForm.Items("Grid01").Specific
    
    oForm.DataSources.DataTables.Add ("PH_PY008")
    
    oGrid1.DataTable = oForm.DataSources.DataTables.Item("PH_PY008")
    Set oDS_PH_PY008 = oForm.DataSources.DataTables.Item("PH_PY008")
    oGrid1.SelectionMode = ms_Auto

    '//----------------------------------------------------------------------------------------------
    '// 기본사항
    '//----------------------------------------------------------------------------------------------
    
    
    oForm.Items("SCLTCOD").DisplayDesc = True
    
    '//일자_조회
    Call oForm.DataSources.UserDataSources.Add("SPosDate", dt_DATE, 8)
    oForm.Items("SPosDate").Specific.DataBind.SetBound True, "", "SPosDate"
    
    
    Call oForm.DataSources.UserDataSources.Add("MSTCOD", dt_SHORT_TEXT, 10)
    oForm.Items("MSTCOD").Specific.DataBind.SetBound True, "", "MSTCOD"
    
    Call oForm.DataSources.UserDataSources.Add("FullName", dt_SHORT_TEXT, 10)
    oForm.Items("FullName").Specific.DataBind.SetBound True, "", "FullName"
    '//일자
    Call oForm.DataSources.UserDataSources.Add("PosDate", dt_DATE, 8)
    oForm.Items("PosDate").Specific.DataBind.SetBound True, "", "PosDate"
    
    
    '//출근
    Call oForm.DataSources.UserDataSources.Add("GetDate", dt_DATE, 8)
    oForm.Items("GetDate").Specific.DataBind.SetBound True, "", "GetDate"
    
    '//퇴근
    Call oForm.DataSources.UserDataSources.Add("OffDate", dt_DATE, 8)
    oForm.Items("OffDate").Specific.DataBind.SetBound True, "", "OffDate"
   
    
    Call oForm.DataSources.UserDataSources.Add("Rotation", dt_SHORT_TEXT, 4)
    oForm.Items("Rotation").Specific.DataBind.SetBound True, "", "Rotation"
    
    Call oForm.DataSources.UserDataSources.Add("GetTime", dt_SHORT_TEXT, 4)
    oForm.Items("GetTime").Specific.DataBind.SetBound True, "", "GetTime"
    
    Call oForm.DataSources.UserDataSources.Add("OffTime", dt_SHORT_TEXT, 4)
    oForm.Items("OffTime").Specific.DataBind.SetBound True, "", "OffTime"
    
    '//기찰기 출근
    Call oForm.DataSources.UserDataSources.Add("RToDate", dt_DATE, 8)
    oForm.Items("RToDate").Specific.DataBind.SetBound True, "", "RToDate"
    
    Call oForm.DataSources.UserDataSources.Add("RToTime", dt_SHORT_TEXT, 4)
    oForm.Items("RToTime").Specific.DataBind.SetBound True, "", "RToTime"
    
    
    '//기찰기 퇴근
    Call oForm.DataSources.UserDataSources.Add("ROffDate", dt_DATE, 8)
    oForm.Items("ROffDate").Specific.DataBind.SetBound True, "", "ROffDate"
    
    Call oForm.DataSources.UserDataSources.Add("ROffTime", dt_SHORT_TEXT, 4)
    oForm.Items("ROffTime").Specific.DataBind.SetBound True, "", "ROffTime"
    
    
    Call oForm.DataSources.UserDataSources.Add("ROutTime", dt_SHORT_TEXT, 4)
    oForm.Items("ROutTime").Specific.DataBind.SetBound True, "", "ROutTime"
    
    Call oForm.DataSources.UserDataSources.Add("RInTime", dt_SHORT_TEXT, 4)
    oForm.Items("RInTime").Specific.DataBind.SetBound True, "", "RInTime"
    
    '//근무형태
    Call oForm.DataSources.UserDataSources.Add("SShiftDat", dt_SHORT_TEXT, 10)
    oForm.Items("SShiftDat").Specific.DataBind.SetBound True, "", "SShiftDat"
    
    
    '//근무조
    Call oForm.DataSources.UserDataSources.Add("SGNMUJO", dt_SHORT_TEXT, 10)
    oForm.Items("SGNMUJO").Specific.DataBind.SetBound True, "", "SGNMUJO"
    
    '//직무코드
    Call oForm.DataSources.UserDataSources.Add("ActCode", dt_SHORT_TEXT, 10)
    oForm.Items("ActCode").Specific.DataBind.SetBound True, "", "ActCode"
    
    Call oForm.DataSources.UserDataSources.Add("ActText", dt_SHORT_TEXT, 50)
    oForm.Items("ActText").Specific.DataBind.SetBound True, "", "ActText"
    
    '//기본시간
    Call oForm.DataSources.UserDataSources.Add("Base", dt_QUANTITY, 10)
    oForm.Items("Base").Specific.DataBind.SetBound True, "", "Base"
    
    '//연장시간
    Call oForm.DataSources.UserDataSources.Add("Extend", dt_QUANTITY, 10)
    oForm.Items("Extend").Specific.DataBind.SetBound True, "", "Extend"
    
    '//특근
    Call oForm.DataSources.UserDataSources.Add("Special", dt_QUANTITY, 10)
    oForm.Items("Special").Specific.DataBind.SetBound True, "", "Special"
    
    '//특근연장
    Call oForm.DataSources.UserDataSources.Add("SpExtend", dt_QUANTITY, 10)
    oForm.Items("SpExtend").Specific.DataBind.SetBound True, "", "SpExtend"
    
    '//심야시간
    Call oForm.DataSources.UserDataSources.Add("Midnight", dt_QUANTITY, 10)
    oForm.Items("Midnight").Specific.DataBind.SetBound True, "", "Midnight"
    
    '//지각시간
    Call oForm.DataSources.UserDataSources.Add("LateTo", dt_QUANTITY, 10)
    oForm.Items("LateTo").Specific.DataBind.SetBound True, "", "LateTo"
    
    Call oForm.DataSources.UserDataSources.Add("EarlyOff", dt_QUANTITY, 10)
    oForm.Items("EarlyOff").Specific.DataBind.SetBound True, "", "EarlyOff"
    
    Call oForm.DataSources.UserDataSources.Add("EarlyTo", dt_QUANTITY, 10)
    oForm.Items("EarlyTo").Specific.DataBind.SetBound True, "", "EarlyTo"
    
    Call oForm.DataSources.UserDataSources.Add("SEarlyTo", dt_QUANTITY, 10)
    oForm.Items("SEarlyTo").Specific.DataBind.SetBound True, "", "SEarlyTo"
    
    
    Call oForm.DataSources.UserDataSources.Add("EducTran", dt_QUANTITY, 10)
    oForm.Items("EducTran").Specific.DataBind.SetBound True, "", "EducTran"
    
    Call oForm.DataSources.UserDataSources.Add("GoOut", dt_QUANTITY, 10)
    oForm.Items("GoOut").Specific.DataBind.SetBound True, "", "GoOut"
    
    Call oForm.DataSources.UserDataSources.Add("SDay", dt_SHORT_TEXT, 10)
    oForm.Items("SDay").Specific.DataBind.SetBound True, "", "SDay"
    
    Call oForm.DataSources.UserDataSources.Add("Day1", dt_SHORT_TEXT, 10)
    oForm.Items("Day1").Specific.DataBind.SetBound True, "", "Day1"
    
    Call oForm.DataSources.UserDataSources.Add("Day2", dt_SHORT_TEXT, 10)
    oForm.Items("Day2").Specific.DataBind.SetBound True, "", "Day2"
    
    Call oForm.DataSources.UserDataSources.Add("GOFrTime", dt_SHORT_TEXT, 4)
    oForm.Items("GOFrTime").Specific.DataBind.SetBound True, "", "GOFrTime"
    
    Call oForm.DataSources.UserDataSources.Add("GOToTime", dt_SHORT_TEXT, 4)
    oForm.Items("GOToTime").Specific.DataBind.SetBound True, "", "GOToTime"
    
    Call oForm.DataSources.UserDataSources.Add("GOFrTim2", dt_SHORT_TEXT, 4)
    oForm.Items("GOFrTim2").Specific.DataBind.SetBound True, "", "GOFrTim2"
    
    Call oForm.DataSources.UserDataSources.Add("GOToTim2", dt_SHORT_TEXT, 4)
    oForm.Items("GOToTim2").Specific.DataBind.SetBound True, "", "GOToTim2"
    
    '비고(2013.11.19 송명규 추가)
    Call oForm.DataSources.UserDataSources.Add("Comment", dt_SHORT_TEXT, 30)
    oForm.Items("Comment").Specific.DataBind.SetBound True, "", "Comment"
    
    Dim oChkBox As SAPbouiCOM.CheckBox
    Call oForm.DataSources.UserDataSources.Add("Chk", dt_SHORT_TEXT, 1)
    Set oChkBox = oForm.Items("Chk").Specific
    oChkBox.ValOn = "Y"
    oChkBox.ValOff = "N"
    oChkBox.DataBind.SetBound True, , "Chk"
    oForm.DataSources.UserDataSources("Chk").VALUE = "N"
    
'    '//부서
'    Set oCombo = oForm.Items("STeamCode").Specific
'
'    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
'                        sQry = sQry & " WHERE Code = '1' AND U_Char2 = '" & oForm.Items.Item("SCLTCOD").Specific.Value & "'"
'                        sQry = sQry & " ORDER BY U_Seq"
'    Call SetReDataCombo(oForm, sQry, oCombo, "Y")
'
'    oForm.Items("STeamCode").DisplayDesc = True
'
'    '//부서
'    Set oCombo = oForm.Items("TeamCode").Specific
'
'    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
'                        sQry = sQry & " WHERE Code = '1' AND U_Char2 = '" & oForm.Items.Item("SCLTCOD").Specific.Value & "'"
'                        sQry = sQry & " ORDER BY U_Seq"
'    Call SetReDataCombo(oForm, sQry, oCombo, "Y")
'
'    oForm.Items("TeamCode").DisplayDesc = True
    
    
    '//담당
    oForm.Items("SRspCode").DisplayDesc = True
    
    
    '// 근무형태
    Set oCombo = oForm.Items("ShiftDat").Specific
    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P154' AND U_UseYN= 'Y' "
    Call SetReDataCombo(oForm, sQry, oCombo, "Y")
    oForm.Items("ShiftDat").DisplayDesc = True

    '// 근무조
    Set oCombo = oForm.Items("GNMUJO").Specific
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P155' AND U_UseYN= 'Y'"
    Call SetReDataCombo(oForm, sQry, oCombo, "Y")
    oForm.Items("GNMUJO").DisplayDesc = True
    
    '//근태
    Set oCombo = oForm.Items("WorkType").Specific
    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P221' AND U_UseYN= 'Y' Order by U_Code "
    Call SetReDataCombo(oForm, sQry, oCombo, "Y")
    oForm.Items("WorkType").DisplayDesc = True
    
    '//위해
    Set oCombo = oForm.Items("DangerCD").Specific
    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P220' And U_Char2 = '" & oForm.Items("SCLTCOD").Specific.VALUE & "' AND U_UseYN= 'Y' "
    Call SetReDataCombo(oForm, sQry, oCombo, "Y")
    oForm.Items("DangerCD").DisplayDesc = True
'    Set oCombo = oForm.Items("WorkType").Specific
    
    
    '// 휴일
    Set oCombo = oForm.Items("SDayOff").Specific
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P202' AND U_UseYN= 'Y'"
    Call SetReDataCombo(oForm, sQry, oCombo)
    oForm.Items("SDayOff").DisplayDesc = True
    
    '// 휴일
    Set oCombo = oForm.Items("DayOff1").Specific
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P202' AND U_UseYN= 'Y'"
    Call SetReDataCombo(oForm, sQry, oCombo)
    oForm.Items("DayOff1").DisplayDesc = True
    
    '// 휴일
    Set oCombo = oForm.Items("DayOff2").Specific
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P202' AND U_UseYN= 'Y'"
    Call SetReDataCombo(oForm, sQry, oCombo)
    oForm.Items("DayOff2").DisplayDesc = True
    
'    '// 근태이상자 확인
    Set oCombo = oForm.Items("Confirm").Specific
    oCombo.ValidValues.Add "N", "미확인[N]"
    oCombo.ValidValues.Add "Y", "확인[Y]"
    oForm.Items("Confirm").DisplayDesc = True
    
    
    '// 교대근무인정
    Set oCombo = oForm.Items("RotateYN").Specific
    oCombo.ValidValues.Add "Y", "인정[Y]"
    oCombo.ValidValues.Add "N", "미인정[N]"
    oForm.Items("RotateYN").DisplayDesc = True
    
    
'    oForm.Items("SPosDate").Specific.Value = Format(Now, "yyyymmdd")

    '1210 창원관리담당, 2210 부산 관리담당, 3210 포장사업장 지원담당
    sQry = "Select TeamCode = U_TeamCode, RspCode = Isnull(U_RspCode,'') "
    sQry = sQry + " From [@PH_PY001A] Where Code = '" & MDC_PS_Common.User_MSTCOD & "'"
    oRecordSet.DoQuery sQry
    
    If Right(Trim(oRecordSet.Fields(0).VALUE), 3) = "200" And Trim(oRecordSet.Fields(1).VALUE) = "" Then
        '운영지원팀
        oRspCodeYN = "Y"
    Else
        If Right(Trim(oRecordSet.Fields(1).VALUE), 3) = "210" Then
            '관리담당
            oRspCodeYN = "Y"
        Else
            oRspCodeYN = "N"
        End If
    End If
    
    
    '관리담당이면 현재일자 이전및 당일 16시 이후에 수정삭제 가능. 그외는 제한을 둠
                    
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
    
PH_PY008_CreateItems_Error:
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY008_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY008_EnableMenus()
    
    On Error GoTo PH_PY008_EnableMenus_Error

    Call oForm.EnableMenu("1283", False)     '//제거
    Call oForm.EnableMenu("1284", False)    '//취소
    Call oForm.EnableMenu("1293", False)     '//행삭제
    
    Exit Sub
    
PH_PY008_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY008_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY008_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY008_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY008_FormItemEnabled
        Call PH_PY008_AddMatrixRow
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY008_FormItemEnabled
        oForm.Items("Code").Specific.VALUE = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY008_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY008_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY008_FormItemEnabled()
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim sQry        As String
    Dim i           As Long
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    
    Dim CLTCOD As String
    Dim sPosDate As String
    
    On Error GoTo PH_PY008_FormItemEnabled_Error
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Call oForm.Freeze(True)
    If (oForm.Mode = fm_ADD_MODE) Then
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "SCLTCOD")
        
        Set oCombo = oForm.Items("SCLTCOD").Specific
        CLTCOD = MDC_SetMod.Get_ReData("Branch", "USER_CODE", "OUSR", "'" & oCompany.UserName & "'")
        oCombo.Select CLTCOD, psk_ByValue
        
        oForm.Items.Item("MSTCOD").Enabled = True
        oForm.Items.Item("FullName").Enabled = False
        oForm.Items.Item("TeamCode").Enabled = True
        oForm.Items.Item("RspCode").Enabled = True
        oForm.Items.Item("SCLTCOD").Enabled = True
        
        'oForm.Items("STeamCode").CLICK ct_Regular
        
        oForm.Items("SPosDate").Specific.VALUE = Format(Now, "yyyymmdd")
        
        oForm.Items("MSTCOD").Specific.VALUE = ""
        oForm.Items("FullName").Specific.VALUE = ""
        
        Set oCombo = oForm.Items("STeamCode").Specific
        
        If oCombo.ValidValues.Count > 0 Then
            For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                oCombo.ValidValues.Remove i, psk_Index
            Next i
            oCombo.ValidValues.Add "", ""
            
            sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
            sQry = sQry & " WHERE Code = '1' AND U_Char2 = '" & oForm.Items.Item("SCLTCOD").Specific.VALUE & "'"
            sQry = sQry & " AND U_UseYN = 'Y'"
            sQry = sQry & " Order By U_Seq"
            Call SetReDataCombo(oForm, sQry, oCombo, "Y")
        End If
        oCombo.Select "", psk_ByValue
        
        '//부서
        Set oCombo = oForm.Items("TeamCode").Specific
        
        If oCombo.ValidValues.Count > 0 Then
            For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                oCombo.ValidValues.Remove i, psk_Index
            Next i
            oCombo.ValidValues.Add "", ""
            
            sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
            sQry = sQry & " WHERE Code = '1' AND U_Char2 = '" & oForm.Items.Item("SCLTCOD").Specific.VALUE & "'"
            sQry = sQry & " AND U_UseYN = 'Y'"
            sQry = sQry & " Order By U_Seq"
            Call SetReDataCombo(oForm, sQry, oCombo, "Y")
        End If
        
        
        
        oCombo.Select "", psk_ByValue
        
        '//담당
        Set oCombo = oForm.Items("RspCode").Specific
        oCombo.Select "", psk_ByValue
        
        '//반
        Set oCombo = oForm.Items("ClsCode").Specific
        oCombo.Select "", psk_ByValue
        
        Set oCombo = oForm.Items("ShiftDat").Specific
        oCombo.Select "", psk_ByValue
        
        Set oCombo = oForm.Items("GNMUJO").Specific
        oCombo.Select "", psk_ByValue
        
        oForm.Items("SShiftDat").Specific.VALUE = ""
        oForm.Items("ShiftDatNm").Specific.VALUE = ""
        oForm.Items("SGNMUJO").Specific.VALUE = ""
        oForm.Items("SGNMUJONm").Specific.VALUE = ""
        
        oForm.Items("ActCode").Specific.VALUE = ""
        oForm.Items("ActText").Specific.VALUE = ""
        
        oForm.DataSources.UserDataSources.Item("RToTime").VALUE = "0000"
        oForm.DataSources.UserDataSources.Item("ROffTime").VALUE = "0000"
        oForm.DataSources.UserDataSources.Item("ROutTime").VALUE = "0000"
        oForm.DataSources.UserDataSources.Item("RInTime").VALUE = "0000"
        
        oForm.DataSources.UserDataSources.Item("RToDate").VALUE = ""
        oForm.DataSources.UserDataSources.Item("ROffDate").VALUE = ""
        
        oForm.DataSources.UserDataSources.Item("Comment").VALUE = "" '비고 추가(2013.11.19 송명규)
        
        CLTCOD = oForm.Items("SCLTCOD").Specific.VALUE
        sPosDate = oForm.Items("SPosDate").Specific.VALUE
        
        sQry = "Select U_WorkType From [@PH_PY003A] a Inner Join [@PH_PY003B] b On a.Code = b.Code "
        sQry = sQry + " Where a.U_CLTCOD = '" & CLTCOD & "' And b.U_Date = '" & sPosDate & "'"
        oRecordSet.DoQuery sQry
        
        Set oCombo = oForm.Items("WorkType").Specific
        If oRecordSet.RecordCount = 0 Then
            Call MDC_Com.MDC_GF_Message("근태월력의 근무일 등록을 하지않습니다.", "E")
            oCombo.Select "", psk_ByValue
        Else
            oCombo.Select PH_PY008_WorkTypeSelect(oForm.Items("SPosDate").Specific.VALUE), psk_ByValue
        End If
    
        
'        oCombo.Select PH_PY008_WorkTypeSelect(Format(Now, "yyyymmdd")), psk_ByValue
'        oCombo.Select "", psk_ByValue
        
'        '//시간reset
'        Call PH_PY008_Time_ReSet
        
        
        Call oForm.EnableMenu("1281", False)     '//문서찾기
        Call oForm.EnableMenu("1282", True)    '//문서추가

    ElseIf (oForm.Mode = fm_FIND_MODE) Then
        oForm.Items.Item("FullName").Enabled = True
        oForm.Items.Item("CLTCOD").Enabled = True
        
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "SCLTCOD")
        
        Call oForm.EnableMenu("1281", False)    '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가
        
        
    ElseIf (oForm.Mode = fm_OK_MODE) Then
        oForm.Items.Item("FullName").Enabled = False
        oForm.Items.Item("CLTCOD").Enabled = False
        
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD", False)
        
        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가
        
    End If
    Set oCombo = Nothing
    Call oForm.Freeze(False)
    Exit Sub
    
PH_PY008_FormItemEnabled_Error:
    Set oCombo = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY008_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Public Sub Raise_FormItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
    Dim sQry        As String
    Dim i           As Long
    Dim Chkcnt      As Long
    Dim tSex        As String
    Dim tBrith      As String
    Dim Day         As String
    Dim ActCode     As String
    Dim MSTCOD      As String
    Dim CLTCOD      As String
    Dim YY          As String
    Dim ymd         As String
    Dim JanQty      As Double
    
    
    
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns     As SAPbouiCOM.Columns
    Dim oRecordSet  As SAPbobsCOM.Recordset
    Dim oRecordSet01  As SAPbobsCOM.Recordset
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Dim ErrNum As Integer
    
    On Error GoTo Raise_FormItemEvent_Error
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Select Case pval.EventType
        Case et_ITEM_PRESSED: '//1
            
            If pval.BeforeAction = True Then
                If pval.ItemUID = "1" Then
                    If PH_PY008_DataValidCheck = False Then
                        BubbleEvent = False
                    End If
                End If
                
                If pval.ItemUID = "Btn_ret" Then
                    Call oForm.Freeze(True)

                    oForm.DataSources.UserDataSources("Chk").VALUE = "N"

'                    Call PH_PY008_ChDate_Set
                    Set oCombo = oForm.Items("WorkType").Specific
                    oCombo.Select oWorkType, psk_ByValue
                    Call PH_PY008_Time_ReSet

                    Call oForm.Freeze(False)
                    
                    Call PH_PY008_MTX01("A")
                End If
                
                
                
                If pval.ItemUID = "Btn01" Then
                    If PH_PY008_ModifyYN() = "Y" Then
                        Call PH_PY008_SAVE
                    End If
                    
                End If
                
                If pval.ItemUID = "Btn02" Then
                    oForm.DataSources.UserDataSources("Chk").VALUE = "Y"
                    Call PH_PY008_MTX01("S") '//근태이상자
                End If
                
                If pval.ItemUID = "Btn_del" Then
                    If PH_PY008_ModifyYN() = "Y" Then
                        Call PH_PY008_Delete
                    End If
                End If
                
                If pval.ItemUID = "Btn_Chk" Then
                    Call oForm.Freeze(True)
                    If oDS_PH_PY008.Rows.Count >= 0 Then
                        For i = 0 To oDS_PH_PY008.Rows.Count - 1
                            If oDS_PH_PY008.Columns.Item("Chk").Cells(i).VALUE = "Y" Then
                                oDS_PH_PY008.Columns.Item("Chk").Cells(i).VALUE = "N"
                            Else
                                oDS_PH_PY008.Columns.Item("Chk").Cells(i).VALUE = "Y"
                            End If
                        Next
                        oForm.Items("Btn_Chk").Specific.Caption = "선택"
                    End If
                    Call oForm.Freeze(False)
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
                            Call PH_PY008_FormItemEnabled
                            Call PH_PY008_AddMatrixRow
                        End If
                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY008_FormItemEnabled
                            Call PH_PY008_AddMatrixRow
                        End If
                    ElseIf oForm.Mode = fm_OK_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY008_FormItemEnabled
                        End If
                    End If
                
'                Case "PIC02"
'                    Call PH_PY008_LoadPic
                End Select
            End If
     
'----------------------------------------------------------
        Case et_KEY_DOWN: '//2
            If pval.BeforeAction = True Then
                If pval.CharPressed = 9 Then
                    If pval.ItemUID = "SShiftDat" Then
                        If oForm.Items("SShiftDat").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    ElseIf pval.ItemUID = "SGNMUJO" Then
                        If oForm.Items("SGNMUJO").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    ElseIf pval.ItemUID = "SMSTCOD" Then
                        If oForm.Items("SMSTCOD").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    ElseIf pval.ItemUID = "MSTCOD" Then
                        If oForm.Items("MSTCOD").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    ElseIf pval.ItemUID = "ActCode" Then
                        If oForm.Items("ActCode").Specific.VALUE = "" Then
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
                    Select Case pval.ItemUID
                        Case "SCLTCOD"
                        
                            Set oCombo = oForm.Items("STeamCode").Specific
        
                            If oCombo.ValidValues.Count > 0 Then
                                For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                                    oCombo.ValidValues.Remove i, psk_Index
                                Next i
                                oCombo.ValidValues.Add "", ""
                                oCombo.Select 0, psk_Index
                            End If
                            
                            sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
                            sQry = sQry & " WHERE Code = '1' And U_Char2 = '" & oForm.Items.Item("SCLTCOD").Specific.VALUE & "' AND U_Char2 = '" & oForm.Items.Item("SCLTCOD").Specific.VALUE & "'"
                            sQry = sQry & " And U_UseYN = 'Y'"
                            sQry = sQry & " Order By U_Seq"
                            Call SetReDataCombo(oForm, sQry, oCombo, "Y")
                            
                            oForm.Items("STeamCode").DisplayDesc = True
                            
                            
                            Set oCombo = oForm.Items("TeamCode").Specific
        
                            If oCombo.ValidValues.Count > 0 Then
                                For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                                    oCombo.ValidValues.Remove i, psk_Index
                                Next i
                                oCombo.ValidValues.Add "", ""
                                oCombo.Select 0, psk_Index
                            End If
                            
                            sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
                            sQry = sQry & " WHERE Code = '1' And U_Char2 = '" & oForm.Items.Item("SCLTCOD").Specific.VALUE & "' AND U_Char2 = '" & oForm.Items.Item("SCLTCOD").Specific.VALUE & "'"
                            sQry = sQry & " And U_UseYN = 'Y'"
                            sQry = sQry & " Order By U_Seq"
                            Call SetReDataCombo(oForm, sQry, oCombo, "Y")
                            
                            oForm.Items("TeamCode").DisplayDesc = True
                            
                            '//위해코드
                            Set oCombo = oForm.Items("DangerCD").Specific
        
                            If oCombo.ValidValues.Count > 0 Then
                                For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                                    oCombo.ValidValues.Remove i, psk_Index
                                Next i
                                oCombo.ValidValues.Add "", ""
                                oCombo.Select 0, psk_Index
                            End If
                            
                            sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
                            sQry = sQry & " WHERE Code = 'P220' And U_Char2 = '" & oForm.Items.Item("SCLTCOD").Specific.VALUE & "' AND U_Char2 = '" & oForm.Items.Item("SCLTCOD").Specific.VALUE & "'"
                            sQry = sQry & " And U_UseYN = 'Y'"
                            sQry = sQry & " Order By U_Seq"
                            Call SetReDataCombo(oForm, sQry, oCombo, "Y")
                            
                            oForm.Items("DangerCD").DisplayDesc = True
                            
                        
                        Case "STeamCode"
                            
                            '//담당 (부서에 따른 담당변경)
                            
                            Set oCombo = oForm.Items("SRspCode").Specific
        
                            If oCombo.ValidValues.Count > 0 Then
                                For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                                    oCombo.ValidValues.Remove i, psk_Index
                                Next i
                                oCombo.ValidValues.Add "", ""
                                oCombo.Select 0, psk_Index
                            End If
                            
                            sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
                            sQry = sQry & " WHERE Code = '2' And U_Char1 = '" & oForm.Items.Item("STeamCode").Specific.VALUE & "' AND U_Char2 = '" & oForm.Items.Item("SCLTCOD").Specific.VALUE & "'"
                            sQry = sQry & " And U_UseYN = 'Y'"
                            sQry = sQry & " Order By U_Seq"
                            Call SetReDataCombo(oForm, sQry, oCombo, "Y")
                            
                            oForm.Items("SRspCode").DisplayDesc = True
                        Case "TeamCode"
                        
                            
                            '//담당 (부서에 따른 담당변경)
                            
                            Set oCombo = oForm.Items("RspCode").Specific
        
                            If oCombo.ValidValues.Count > 0 Then
                                For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                                    oCombo.ValidValues.Remove i, psk_Index
                                Next i
                                oCombo.ValidValues.Add "", ""
                                oCombo.Select 0, psk_Index
                            End If
                            
                            sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
                            sQry = sQry & " WHERE Code = '2' And U_Char1 = '" & oForm.Items.Item("TeamCode").Specific.VALUE & "' AND U_Char2 = '" & oForm.Items.Item("SCLTCOD").Specific.VALUE & "'"
                            sQry = sQry & " And U_UseYN = 'Y'"
                            sQry = sQry & " Order By U_Seq"
                            Call SetReDataCombo(oForm, sQry, oCombo, "Y")
                            
                            oForm.Items("RspCode").DisplayDesc = True
                        Case "SRspCode"
                            
                            '//반변경 (담당에 따른 반변경)
                            
                            Set oCombo = oForm.Items("SClsCode").Specific
        
                            If oCombo.ValidValues.Count > 0 Then
                                For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                                    oCombo.ValidValues.Remove i, psk_Index
                                Next i
                                oCombo.ValidValues.Add "", ""
                                oCombo.Select 0, psk_Index
                            End If
                            
                            sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
                            sQry = sQry & " WHERE Code = '9' And U_Char1 = '" & oForm.Items.Item("SRspCode").Specific.VALUE & "' AND U_Char3 = '" & oForm.Items.Item("SCLTCOD").Specific.VALUE & "'"
                            sQry = sQry & " And U_UseYN = 'Y'"
                            sQry = sQry & " Order By U_Seq"
                            Call SetReDataCombo(oForm, sQry, oCombo, "Y")
                            
                            oForm.Items("SClsCode").DisplayDesc = True
                        Case "RspCode"
                        
                            
                            '//반변경 (담당에 따른 반변경)
                            
                            Set oCombo = oForm.Items("ClsCode").Specific
        
                            If oCombo.ValidValues.Count > 0 Then
                                For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                                    oCombo.ValidValues.Remove i, psk_Index
                                Next i
                                oCombo.ValidValues.Add "", ""
                                oCombo.Select 0, psk_Index
                            End If
                            
                            sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
                            sQry = sQry & " WHERE Code = '9' And U_Char1 = '" & oForm.Items.Item("RspCode").Specific.VALUE & "' AND U_Char3 = '" & oForm.Items.Item("SCLTCOD").Specific.VALUE & "'"
                            sQry = sQry & " And U_UseYN = 'Y'"
                            sQry = sQry & " Order By U_Seq"
                            Call SetReDataCombo(oForm, sQry, oCombo, "Y")
                            
                            oForm.Items("ClsCode").DisplayDesc = True
                        Case "WorkType"
                        
                            Select Case oForm.Items("WorkType").Specific.VALUE
                                Case "A00"
                                    '//정상근무
                                    oForm.Items("Rotation").Specific.VALUE = 1
                                Case "A01", "A02", "E02", "E03", "F01", "F02", "F03", "F04", "F05", "D11"
                                    '//무단결근, 유계결근, 무급휴일, 휴업, 공상휴직, 병가(휴직), 신병휴직, 정직(유결), 가사휴직
                                    oForm.Items("OffDate").Specific.VALUE = oForm.Items("GetDate").Specific.VALUE
                                    oForm.Items("GetTime").Specific.VALUE = "0000"
                                    oForm.Items("OffTime").Specific.VALUE = "0000"
                                    
                                    Call PH_PY008_Time_ReSet
                                    
                                    oForm.Items("Rotation").Specific.VALUE = 0
                                Case "C02", "D04", "D05", "D06", "D07", "H05"
                                    '//훈련, 경조휴가, 하기휴가, 특별휴가, 분만휴가, 조합활동
                                    
                                    oForm.Items("GetTime").Specific.VALUE = "0000"
                                    oForm.Items("OffTime").Specific.VALUE = "0000"
                                    Call PH_PY008_Time_ReSet
                                    oForm.Items("OffDate").Specific.VALUE = oForm.Items("GetDate").Specific.VALUE
                                    
                                    oForm.Items("Rotation").Specific.VALUE = 0
                                Case "D02", "D09"
                                    '//연차/반차 휴가
                                    '//연차/반차 휴가 잔여일수 확인
                                    
                                    If oForm.Items("WorkType").Specific.VALUE = "D02" Then
                                        JanQty = 1
                                    ElseIf oForm.Items("WorkType").Specific.VALUE = "D09" Then
                                        JanQty = 0.5
                                    End If
                                    
                                    ymd = oForm.Items("PosDate").Specific.VALUE
                                    CLTCOD = oForm.Items("SCLTCOD").Specific.VALUE
                                    MSTCOD = oForm.Items("MSTCOD").Specific.VALUE
                                    
                                    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
                                    
                                    If Trim(MSTCOD) <> "" Then
                                        sQry = "exec [PH_PY775_01] '" & CLTCOD & "','"
                                        sQry = sQry & Left(ymd, 4) & "','" & MSTCOD & "'"
                                        oRecordSet.DoQuery sQry
                                        
                                        If oRecordSet.Fields("jandd").VALUE < JanQty Then
                                        
                                            ErrNum = 1
                                            
                                            Set oCombo = oForm.Items("WorkType").Specific
                                            Call oCombo.Select("A00", psk_ByValue)
                                            
                                            GoTo Raise_FormItemEvent_Error
                                            
                                        Else
                                            oForm.Items("GetTime").Specific.VALUE = "0000"
                                            oForm.Items("OffTime").Specific.VALUE = "0000"
                                            Call PH_PY008_Time_ReSet
                                            oForm.Items("OffDate").Specific.VALUE = oForm.Items("GetDate").Specific.VALUE
                                            
                                            oForm.Items("Rotation").Specific.VALUE = 1
                                        End If
                                    Else
                                        Sbo_Application.SetStatusBarMessage "사번정보가 없습니다. 확인바랍니다.", bmt_Short, True
                                            
                                        Set oCombo = oForm.Items("WorkType").Specific
                                        oCombo.Select "A00", psk_ByValue
                                    End If
                                    
                                    ProgBar01.VALUE = 100
                                    Call ProgBar01.Stop
                                    
                                Case "D08", "D10"
                                    '//근속보전휴가, 근속보전반차(기계사업부)
                                    '//근속보전휴가 잔량 확인
                                    ymd = oForm.Items("PosDate").Specific.VALUE
                                    CLTCOD = oForm.Items("SCLTCOD").Specific.VALUE
                                    MSTCOD = oForm.Items("MSTCOD").Specific.VALUE
                                    
                                    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
                                    
                                    If oForm.Items("PosDate").Specific.VALUE >= Left(ymd, 4) + "0701" Then
                                        YY = Left(ymd, 4)
                                    Else
                                        YY = CStr(CInt(Left(ymd, 4)) - 1)
                                    End If
                                    
'                                    sQry = "Select Bqty     = Isnull( ( SELECT Isnull(c.Qty,0) "
'                                    sQry = sQry + " FROM [ZPH_GUNSOKCHA] c"
'                                    sQry = sQry + " WHERE c.Year = '" & YY & "'"
'                                    sQry = sQry + " AND (CASE WHEN ISDATE(a.U_GrpDat) = 0 THEN a.U_StartDat "
'                                    sQry = sQry + " Else A.U_GrpDat END) BETWEEN c.DocDateFr AND c.DocDateTo ),0), " '발생수량
'                                    sQry = sQry + " Sqty     = Isnull( ( SELECT Count(*)"
'                                    sQry = sQry + " FROM [ZPH_PY008] c"
'                                    sQry = sQry + " Where c.CLTCOD = A.U_CLTCOD"
'                                    sQry = sQry + " AND c.PosDate  Between '" & YY & "' + '0701' and convert(char(4),convert(integer,'" & YY & "') + 1 ) + '0630'"
'                                    sQry = sQry + " AND c.MSTCOD  = a.Code"
'                                    sQry = sQry + " AND c.WorkType = 'D08'"
'                                    sQry = sQry + " And c.PosDate <> '" & ymd & "' ),0) + "  '사용수량
'                                    sQry = sQry + " Isnull( ( SELECT Count(*) / 2.0"
'                                    sQry = sQry + " FROM [ZPH_PY008] c"
'                                    sQry = sQry + " Where c.CLTCOD = A.U_CLTCOD"
'                                    sQry = sQry + " AND c.PosDate  Between '" & YY & "' + '0701' and convert(char(4),convert(integer,'" & YY & "') + 1 ) + '0630'"
'                                    sQry = sQry + " AND c.MSTCOD  = a.Code"
'                                    sQry = sQry + " AND c.WorkType = 'D10'"
'                                    sQry = sQry + " And c.PosDate <> '" & ymd & "' ),0)" '근속반차 사용수량
'                                    sQry = sQry + " FROM [@PH_PY001A] a"
'                                    sQry = sQry + " WHERE a.U_CLTCOD   = '" & CLTCOD & "'"
'                                    sQry = sQry + " AND a.Code  = '" & MSTCOD & "'"
                                    sQry = "exec [PH_PY605_99] '" & CLTCOD & "','" & YY & "', '" & ymd & "', '" & MSTCOD & "'"
                                    
                                    oRecordSet.DoQuery sQry
                                    
                                    If oForm.Items("WorkType").Specific.VALUE = "D08" Then
                                        JanQty = 1
                                    ElseIf oForm.Items("WorkType").Specific.VALUE = "D10" Then
                                        JanQty = 0.5
                                    End If
                                    
                                    If oRecordSet.Fields("Bqty").VALUE - oRecordSet.Fields("Sqty").VALUE < JanQty Then
                                    
                                        ErrNum = 2
                                        Set oCombo = oForm.Items("WorkType").Specific
                                        Call oCombo.Select("A00", psk_ByValue)
                                        
                                        GoTo Raise_FormItemEvent_Error
                                        
                                    Else
                                        oForm.Items("GetTime").Specific.VALUE = "0000"
                                        oForm.Items("OffTime").Specific.VALUE = "0000"
                                        Call PH_PY008_Time_ReSet
                                        oForm.Items("OffDate").Specific.VALUE = oForm.Items("GetDate").Specific.VALUE
                                        
                                        oForm.Items("Rotation").Specific.VALUE = 1
                                    End If
                                    
                                    ProgBar01.VALUE = 100
                                    Call ProgBar01.Stop
                                    
                            End Select
                        
                        Case "ShiftDat"
                            '//근무형태에 따른 근무조 값
                        
                            Set oCombo = oForm.Items("GNMUJO").Specific
        
                            If oCombo.ValidValues.Count > 0 Then
                                For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                                    oCombo.ValidValues.Remove i, psk_Index
                                Next i
                            End If
                            
                            sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
                            sQry = sQry & " WHERE Code = 'P155' AND U_Char1 = '" & oForm.Items.Item("ShiftDat").Specific.VALUE & "'"
                            sQry = sQry & " And U_UseYN = 'Y'"
                            sQry = sQry & " ORDER BY U_Code"
                            Call SetReDataCombo(oForm, sQry, oCombo, "Y")
                            
                            oCombo.Select 0, psk_Index
                            oForm.Items("GNMUJO").DisplayDesc = True
                        Case "GNMUJO"
                        
                            '//시간reset
                            Call PH_PY008_Time_ReSet
                            
                            Call PH_PY008_ChGNMUJO_Set
                            
                        End Select
                    
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
                            If oForm.DataSources.UserDataSources("Chk").VALUE = "N" Then
                                If pval.ColUID <> "Chk" Then
                                    PH_PY008_MTX02 pval.ItemUID, pval.Row, pval.ColUID
                                End If
                            Else
                                PH_PY008_MTX02 pval.ItemUID, pval.Row, pval.ColUID
                            End If
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
                Select Case pval.ItemUID
                Case "Grid01"
                    If pval.Row >= 0 Then
                        Select Case pval.ItemUID
                        Case "Grid01"
                            'Call oMat1.SelectRow(pval.Row, True, False)
                            If oForm.DataSources.UserDataSources("Chk").VALUE = "N" Then
                                If pval.ColUID = "Chk" Then
                                    '//체크수
                                    If oDS_PH_PY008.Rows.Count > 0 Then
                                        
                                        For i = 0 To oDS_PH_PY008.Rows.Count - 1
                                            If oDS_PH_PY008.Columns.Item("Chk").Cells(i).VALUE = "Y" Then
                                                Chkcnt = Chkcnt + 1
                                            End If
                                            
                                        Next
                                        oForm.Items("Chkcnt").Specific.VALUE = Chkcnt
                                    End If
                                End If
                            Else
                                For i = 0 To oDS_PH_PY008.Rows.Count - 1
                                    If oDS_PH_PY008.Columns.Item("Chk").Cells(i).VALUE = "Y" Then
                                        oDS_PH_PY008.Columns.Item("Chk").Cells(i).VALUE = "N"
                                    End If
                                Next
                                oDS_PH_PY008.Columns.Item("Chk").Cells(pval.Row).VALUE = "Y"
                                oForm.Items("Chkcnt").Specific.VALUE = 1
                            End If
                        End Select
                        
                    End If
                End Select
            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_DOUBLE_CLICK: '//7
            Call oForm.Freeze(True)
            If pval.BeforeAction = True Then
            Else
                
'                If pval.ItemUid = "Grid01" Then
'                    If pval.ColUID = "Chk" Then
'                        If pval.Row = -1 Then
'                            If oDS_PH_PY008.Rows.Count >= 0 Then
'                                For i = 0 To oDS_PH_PY008.Rows.Count - 1
'                                    If oDS_PH_PY008.Columns.Item("Chk").Cells(i).Value = "Y" Then
'                                        oDS_PH_PY008.Columns.Item("Chk").Cells(i).Value = "N"
'                                    Else
'                                        oDS_PH_PY008.Columns.Item("Chk").Cells(i).Value = "Y"
'                                    End If
'                                Next
'                            End If
'                        End If
'                    End If
'                End If
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
                    Call oForm.Freeze(True)
                    Select Case pval.ItemUID
                        Case "SPosDate"
                            oGrid1.DataTable.Clear
                            oForm.Items("Chkcnt").Specific.VALUE = 0
'                            oMat1.Clear
                        Case "OffDate"
                            oForm.DataSources.UserDataSources.Item("OffTime").VALUE = "0000"
'                            oForm.Items("OffTime").Specific.Value = "0000"
                        Case "GetTime"
                            oForm.DataSources.UserDataSources.Item("OffTime").VALUE = "0000"
'                            oForm.Items("OffTime").Specific.Value = "0000"
                            Call PH_PY008_Time_ReSet
                    End Select
                    Call oForm.Freeze(False)
                End If
                    
            ElseIf pval.BeforeAction = False Then
                If pval.ItemChanged = True Then
                    Select Case pval.ItemUID

                        Case "SPosDate"
                            
                            Call oForm.Freeze(True)
                            Call PH_PY008_ChDate_Set
                            '//시간reset
                            Call PH_PY008_Time_ReSet
                            
                            oForm.Items("SShiftDat").Specific.VALUE = ""
                            Call oForm.Freeze(False)
                            
                            
                        Case "PosDate"
                            oForm.Items("SPosDate").Specific.VALUE = oForm.Items("PosDate").Specific.VALUE
                            'oForm.DataSources.UserDataSources.Item("SPosDate").Value = oForm.Items("PosDate").Specific.Value
'                            oForm.Items("SPosDate").Specific.Value = oForm.Items("PosDate").Specific.Value
                        Case "GetDate"
                            oForm.DataSources.UserDataSources.Item("SPosDate").VALUE = oForm.Items("GetDate").Specific.VALUE
'                            oForm.Items("SPosDate").Specific.Value = oForm.Items("GetDate").Specific.Value
                        Case "OffDate"
'                            oForm.Items("OffTime").Specific.Value = "0000"
                        Case "SShiftDat"   '//근무형태
                            Call oForm.Freeze(True)
                            If Trim(oForm.Items.Item("SShiftDat").Specific.VALUE) <> "" Then
                                sQry = "SELECT U_CodeNm FROM [@PS_HR200L] "
                                sQry = sQry & " WHERE Code = 'P154' AND U_Code = '" & Trim(oForm.Items.Item("SShiftDat").Specific.VALUE) & "'"
                                sQry = sQry & " ORDER BY U_Code"
                                
                                oRecordSet.DoQuery sQry
                                If oRecordSet.RecordCount > 0 Then
                                   oForm.Items("ShiftDatNm").Specific.VALUE = oRecordSet.Fields(0).VALUE
                                   
                                   sQry = "select u_Team1 from [@PH_PY003B] where left(code,1) ='1' and u_date ='" & Trim(oForm.Items.Item("SPosDate").Specific.VALUE) & "'"
                                   oRecordSet.DoQuery sQry
                                   oForm.Items("Team1").Specific.VALUE = oRecordSet.Fields(0).VALUE
                                   
                                   sQry = "select u_Team2 from [@PH_PY003B] where left(code,1) ='1' and u_date ='" & Trim(oForm.Items.Item("SPosDate").Specific.VALUE) & "'"
                                   oRecordSet.DoQuery sQry
                                   oForm.Items("Team2").Specific.VALUE = oRecordSet.Fields(0).VALUE
                                   
                                   
                                   sQry = "select u_Team3 from [@PH_PY003B] where left(code,1) ='1' and u_date ='" & Trim(oForm.Items.Item("SPosDate").Specific.VALUE) & "'"
                                   oRecordSet.DoQuery sQry
                                   oForm.Items("Team3").Specific.VALUE = oRecordSet.Fields(0).VALUE
                                   

                                    
                                sQry = "SELECT U_CodeNm FROM [@PS_HR200L] "
                                sQry = sQry & " WHERE Code = 'P154' AND U_Code = '" & Trim(oForm.Items.Item("SShiftDat").Specific.VALUE) & "'"
                                sQry = sQry & " ORDER BY U_Code"
                                
                                oRecordSet.DoQuery sQry
                                
                                
                                   oForm.Items("SGNMUJO").Specific.VALUE = ""
                                   oForm.Items("SGNMUJONm").Specific.VALUE = ""
                                   
                                   Set oCombo = oForm.Items("ShiftDat").Specific
                                   oCombo.Select oRecordSet.Fields(0).VALUE, psk_ByValue
                                   
                                Else
                                    oForm.Items("ShiftDatNm").Specific.VALUE = ""
                                    Call MDC_Com.MDC_GF_Message("근무형태가 존재하지 않습니다.", "E")
                                End If
                            Else
                                oForm.Items("ShiftDatNm").Specific.VALUE = ""
                                   
                                oForm.Items("SGNMUJO").Specific.VALUE = ""
                                oForm.Items("SGNMUJONm").Specific.VALUE = ""
                                Set oCombo = oForm.Items("ShiftDat").Specific
                                oCombo.Select "", psk_ByValue
                            End If
                            Call oForm.Freeze(False)
                            
                        Case "SGNMUJO"   '//근무조
                            Call oForm.Freeze(True)
                            If oForm.Items("SGNMUJO").Specific.VALUE <> "" Then
                            
                                sQry = "        SELECT   U_CodeNm "
                                sQry = sQry & " FROM     [@PS_HR200L] "
                                sQry = sQry & " WHERE    Code = 'P155' "
                                sQry = sQry & "          AND U_Code = '" & Trim(oForm.Items.Item("SGNMUJO").Specific.VALUE) & "'"
                                sQry = sQry & "          AND U_Char1 = '" & Trim(oForm.Items.Item("SShiftDat").Specific.VALUE) & "'"
                                sQry = sQry & " ORDER BY U_Code"
                                
                                oRecordSet.DoQuery sQry
                                
                                If Trim(oForm.Items.Item("SShiftDat").Specific.VALUE) = "4" Then '3조 2교대 일 경우만 해당 로직 수행_S(2018.07.04 송명규)
                                    ' WORKTYPE 변경 로직 주 52시간일 경우 휴무일경우 E01 주간/야간일 경우 A00을 입력) (주52)
                                    If (oRecordSet.Fields(0).VALUE = "1조") Then
                                        
                                        sQry = "        Select (case when U_Team1 is null then 'E01' else U_WorkType2 end)"
                                        sQry = sQry + " From   [@PH_PY003A] a Inner Join [@PH_PY003B] b On a.Code = b.Code "
                                        sQry = sQry + " Where  a.U_CLTCOD = '" & Trim(oForm.Items.Item("SCLTCOD").Specific.VALUE) & "'"
                                        sQry = sQry + "        And b.U_Date = '" & Trim(oForm.Items.Item("SPosDate").Specific.VALUE) & "'"
                                        
                                        oRecordSet01.DoQuery sQry
                                        
                                    End If
                                    
                                    If (oRecordSet.Fields(0).VALUE = "2조") Then
                                        
                                        sQry = "        Select (case when U_Team2 is null then 'E01' else U_WorkType2 end)"
                                        sQry = sQry + " From   [@PH_PY003A] a Inner Join [@PH_PY003B] b On a.Code = b.Code "
                                        sQry = sQry + " Where  a.U_CLTCOD = '" & Trim(oForm.Items.Item("SCLTCOD").Specific.VALUE) & "'"
                                        sQry = sQry + "        And b.U_Date = '" & Trim(oForm.Items.Item("SPosDate").Specific.VALUE) & "'"
                                        
                                        oRecordSet01.DoQuery sQry
                                        
                                    End If
                                    
                                    If (oRecordSet.Fields(0).VALUE = "3조") Then
                                        
                                        sQry = "        Select (case when U_Team3 is null then 'E01' else U_WorkType2 end)"
                                        sQry = sQry + " From   [@PH_PY003A] a Inner Join [@PH_PY003B] b On a.Code = b.Code "
                                        sQry = sQry + " Where  a.U_CLTCOD = '" & Trim(oForm.Items.Item("SCLTCOD").Specific.VALUE) & "'"
                                        sQry = sQry + "        And b.U_Date = '" & Trim(oForm.Items.Item("SPosDate").Specific.VALUE) & "'"
                                        
                                        oRecordSet01.DoQuery sQry
                                        
                                    End If
                                    
                                    If oRecordSet01.RecordCount = 0 Then
                                        Call MDC_Com.MDC_GF_Message("근태월력의 근무일 등록을 하지않습니다.", "E")
                                       ' GoTo PH_PY008_ChDate_Set_Exit
                                    Else
                                        oWorkType = oRecordSet01.Fields(0).VALUE
                                        Set oCombo = oForm.Items("WorkType").Specific
                                        oCombo.Select oRecordSet01.Fields(0).VALUE, psk_ByValue
                                      '  Call PH_PY008_ChGNMUJO_Set
                                    End If
                                 '3조 2교대 일 경우만 해당 로직 수행_E(2018.07.04 송명규)
                                Else '3조 2교대가 아닌경우는 근태월력의 근태구분1을 선택(2018.07.04 송명규)
                                
                                    sQry = "        SELECT  T1.U_WorkType"
                                    sQry = sQry & " FROM    [@PH_PY003A] AS T0 "
                                    sQry = sQry & "         INNER JOIN "
                                    sQry = sQry & "         [@PH_PY003B] AS T1 "
                                    sQry = sQry & "             ON T0.Code = T1.Code "
                                    sQry = sQry & " Where   T0.U_CLTCOD = '" & Trim(oForm.Items.Item("SCLTCOD").Specific.VALUE) & "'"
                                    sQry = sQry & "         AND T1.U_Date = '" & Trim(oForm.Items.Item("SPosDate").Specific.VALUE) & "'"
                                    
                                    oRecordSet01.DoQuery sQry
                                
                                    Call oForm.Items("WorkType").Specific.Select(oRecordSet01.Fields(0).VALUE, psk_ByValue)
                                
                                End If
                                
                                If oRecordSet.RecordCount > 0 Then
                                    Call oForm.Freeze(True)
                                    
                                    oForm.Items("SGNMUJONm").Specific.VALUE = oRecordSet.Fields(0).VALUE
                                   
                                    Set oCombo = oForm.Items("GNMUJO").Specific
                                    oCombo.Select oRecordSet.Fields(0).VALUE, psk_ByValue
                                    
                                    Call PH_PY008_ChGNMUJO_Set
                                    
                                    Call oForm.Freeze(False)
                                Else
                                    oForm.Items("SGNMUJONm").Specific.VALUE = ""
                                    Call MDC_Com.MDC_GF_Message("근무형태가 존재하지 않습니다.", "E")
                                End If
                            End If
                            Call oForm.Freeze(False)
                        Case "GetTime"
'                            oForm.Items("OffTime").Specific.Value = "0000"
                        Case "OffTime"
                            If oForm.Items("OffTime").Specific.VALUE <> "0000" Then
                                PH_PY008_Time_Calc_Main (oForm.Items("OffTime").Specific.VALUE)
                            End If
                        Case "ActCode"
                            Call oForm.Freeze(True)
                            ActCode = oForm.Items("ActCode").Specific.VALUE
                            sQry = "select b.U_CodeNm, b.U_Char1 from [@PS_HR200H] a Inner Join [@PS_HR200L] b On a.Code = b.Code And a.Code = 'P127' And b.U_Code = '" & ActCode & "'"
                            oRecordSet.DoQuery sQry
                            Set oCombo = oForm.Items("DangerCD").Specific
                            If oRecordSet.RecordCount > 0 Then
                                oForm.Items("ActText").Specific.VALUE = oRecordSet.Fields(0).VALUE
                                oCombo.Select oRecordSet.Fields(1).VALUE, psk_ByValue
                            Else
                                oForm.Items("ActText").Specific.VALUE = ""
                                oCombo.Select "", psk_ByValue
                            End If
                            Call oForm.Freeze(False)
                        Case "SMSTCOD"
                            Call oForm.Freeze(True)
                            MSTCOD = oForm.Items("SMSTCOD").Specific.VALUE
                            sQry = "SELECT U_FullName from [@PH_PY001A] Where Code = '" & MSTCOD & "'"
                            oRecordSet.DoQuery sQry
                            If oRecordSet.RecordCount > 0 Then
                                oForm.Items("SFullName").Specific.VALUE = oRecordSet.Fields(0).VALUE
                            Else
                                oForm.Items("SFullName").Specific.VALUE = ""
                            End If
                            Call oForm.Freeze(False)
                        Case "MSTCOD"
                            Call oForm.Freeze(True)
                            MSTCOD = oForm.Items("MSTCOD").Specific.VALUE
                            sQry = "SELECT U_FullName, U_TeamCode, U_RspCode, U_ShiftDat, U_GNMUJO from [@PH_PY001A] Where Code = '" & MSTCOD & "'"
                            oRecordSet.DoQuery sQry
                            If oRecordSet.RecordCount > 0 Then
                                oForm.Items("FullName").Specific.VALUE = oRecordSet.Fields(0).VALUE
                                Set oCombo = oForm.Items("TeamCode").Specific
                                oCombo.Select oRecordSet.Fields(1).VALUE, psk_ByValue
                                
                                Set oCombo = oForm.Items("RspCode").Specific
                                oCombo.Select oRecordSet.Fields(2).VALUE, psk_ByValue
                                
                                Set oCombo = oForm.Items("ShiftDat").Specific
                                oCombo.Select oRecordSet.Fields(3).VALUE, psk_ByValue
                                
                                Set oCombo = oForm.Items("GNMUJO").Specific
                                oCombo.Select oRecordSet.Fields(4).VALUE, psk_ByValue
                            Else
                                oForm.Items("FullName").Specific.VALUE = ""
                            End If
                            Call oForm.Freeze(False)
                    End Select
                 
                End If
            End If
'            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_MATRIX_LOAD: '//11
            If pval.BeforeAction = True Then
            ElseIf pval.BeforeAction = False Then
'                oMat1.LoadFromDataSource
'                Call PH_PY008_AddMatrixRow
                
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
                Set oDS_PH_PY008 = Nothing
                
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
                oForm.Items.Item("79").Width = oForm.Items.Item("KUKGRD").Left + oForm.Items.Item("KUKGRD").Width - oForm.Items.Item("79").Left + 10
                oForm.Items.Item("79").Height = oForm.Items.Item("80").Height
                
                oForm.Items.Item("77").Width = oForm.Items.Item("BUYN20").Left + oForm.Items.Item("BUYN20").Width - oForm.Items.Item("77").Left + 16
                oForm.Items.Item("77").Height = oForm.Items.Item("78").Height
            
                
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
'                    Call MDC_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY008A", "Code")
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
    Set ProgBar01 = Nothing
    
    Exit Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Raise_FormItemEvent_Error:
    oForm.Freeze (False)
    Set oCombo = Nothing
    Set oRecordSet = Nothing
    Set ProgBar01 = Nothing
    
    If ErrNum = 1 Then
        Call Sbo_Application.SetStatusBarMessage("일반년차잔여일수가 없습니다. 확인바랍니다.", bmt_Short, True)
    ElseIf ErrNum = 2 Then
        Call Sbo_Application.SetStatusBarMessage("보전휴가잔여일수가 없습니다. 확인바랍니다.", bmt_Short, True)
    Else
        Call Sbo_Application.SetStatusBarMessage("Raise_FormItemEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True)
    End If
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
'                Call PH_PY008_FormItemEnabled
        End Select
    ElseIf (pval.BeforeAction = False) Then
        Select Case pval.MenuUID
            Case "1283":
                oForm.Mode = fm_ADD_MODE
                Call PH_PY008_FormItemEnabled
'                Call PH_PY008_AddMatrixRow
            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//문서찾기
                Call PH_PY008_FormItemEnabled
'                Call PH_PY008_AddMatrixRow
                oForm.Items("Code").CLICK ct_Regular
            Case "1282":    '//문서추가
                Call PH_PY008_FormItemEnabled
'                Call PH_PY008_AddMatrixRow
            Case "1288", "1289", "1290", "1291":
                Call PH_PY008_FormItemEnabled
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

Sub PH_PY008_AddMatrixRow()
    Dim oRow    As Long
    
    On Error GoTo PH_PY008_AddMatrixRow_Error
    
    Call oForm.Freeze(True)
    
'    '//[Mat1]
'    oMat1.FlushToDataSource
'    oRow = oMat1.VisualRowCount
'
'    If oMat1.VisualRowCount > 0 Then
'        If Trim(oDS_PH_PY008A.GetValue("U_MSTCOD", oRow - 1)) <> "" Then
'            If oDS_PH_PY008A.Size <= oMat1.VisualRowCount Then
'                oDS_PH_PY008A.InsertRecord (oRow)
'            End If
'            oDS_PH_PY008A.Offset = oRow
'            oDS_PH_PY008A.setValue "U_LineNum", oRow, oRow + 1
'            oMat1.LoadFromDataSource
'        Else
'            oDS_PH_PY008A.Offset = oRow - 1
'            oDS_PH_PY008A.setValue "U_LineNum", oRow - 1, oRow
'            oMat1.LoadFromDataSource
'        End If
'    ElseIf oMat1.VisualRowCount = 0 Then
'        oDS_PH_PY008A.Offset = oRow
'        oDS_PH_PY008A.setValue "U_LineNum", oRow, oRow + 1
'        oMat1.LoadFromDataSource
'    End If

    

    Call oForm.Freeze(False)
    Exit Sub
PH_PY008_AddMatrixRow_Error:

    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY008_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY008_FormClear()
On Error GoTo PH_PY008_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY008'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PH_PY008_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY008_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY008_DataValidCheck() As Boolean
On Error GoTo PH_PY008_DataValidCheck_Error
    PH_PY008_DataValidCheck = False
    Dim i       As Long
    Dim j       As Long
    
    Dim sQry    As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
'    '//----------------------------------------------------------------------------------
'    '//기본사항 탭
'    '//----------------------------------------------------------------------------------
'    If Trim(oDS_PH_PY008A.GetValue("Code", 0)) = "" Then
'        Sbo_Application.SetStatusBarMessage "사원번호는 필수입니다.", bmt_Short, True
'        oForm.Items("Code").CLICK ct_Regular
'        PH_PY008_DataValidCheck = False
'        Exit Function
'    End If
'
'    If oForm.Mode = fm_ADD_MODE Then
'        '//사원번호 중복 체크 (PH_PY008A)
'        sQry = "SELECT Code FROM [@PH_PY008A] WHERE Code = '" & Trim(oDS_PH_PY008A.GetValue("Code", 0)) & "'"
'        oRecordSet.DoQuery sQry
'
'        If oRecordSet.RecordCount > 0 Then
'            Sbo_Application.SetStatusBarMessage "[PH_PY008A] 이미 존재하는 사원번호입니다", bmt_Short, True
'            oForm.Items("Code").CLICK ct_Regular
'            PH_PY008_DataValidCheck = False
'            Exit Function
'        End If
'
'        '//사원번호 중복 체크(OHEM)
'        sQry = "SELECT U_MSTCOD FROM OHEM WHERE U_MSTCOD = '" & Trim(oDS_PH_PY008A.GetValue("Code", 0)) & "'"
'        oRecordSet.DoQuery sQry
'
'        If oRecordSet.RecordCount > 0 Then
'            Sbo_Application.SetStatusBarMessage "[OHEM] 이미 존재하는 사원번호입니다", bmt_Short, True
'            oForm.Items("Code").CLICK ct_Regular
'            PH_PY008_DataValidCheck = False
'            Exit Function
'        End If
'    End If
'    '// Code = Name 동일
'    Call oDS_PH_PY008A.setValue("Name", 0, oDS_PH_PY008A.GetValue("Code", 0))
'
'
'
'    oMat1.FlushToDataSource
'
'    '// Matrix 마지막 행 삭제(DB 저장시)
'    If oDS_PH_PY008A.Size > 1 Then oDS_PH_PY008B.RemoveRecord (oDS_PH_PY008B.Size - 1)
'
'    oMat1.LoadFromDataSource
'
'    PH_PY008_DataValidCheck = True
    
    Exit Function
    
    Set oRecordSet = Nothing
    
    
PH_PY008_DataValidCheck_Error:
    Set oRecordSet = Nothing
    PH_PY008_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY008_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY008_MTX01(ByVal oConfirm As String) '//oConfirm A: 조회, S:근태이상자

    '//메트릭스에 데이터 로드
    
    Dim i       As Long
    Dim Chkcnt  As Long
    Dim sQry    As String
    Dim iRow        As Long
    
    Dim Param01 As String
    Dim Param02 As String
    Dim Param03 As String
    Dim Param04 As String
    Dim Param05 As String
    Dim Param06 As String
    Dim Param07 As String
    Dim Param08 As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY008_MTX01_Error

    Call oForm.Freeze(True)
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)

    Param01 = oForm.Items("SCLTCOD").Specific.VALUE
    Param02 = oForm.Items("SPosDate").Specific.VALUE
    Param03 = oForm.Items("STeamCode").Specific.VALUE
    Param04 = oForm.Items("SRspCode").Specific.VALUE
    Param05 = oForm.Items("SClsCode").Specific.VALUE
    Param06 = oForm.Items("SMSTCOD").Specific.VALUE
    Param07 = oForm.Items("SShiftDat").Specific.VALUE
    Param08 = oForm.Items("SGNMUJO").Specific.VALUE
    
    '//해당일의 근무일 등록 확인
    sQry = "Select U_WorkType From [@PH_PY003A] a Inner Join [@PH_PY003B] b On a.Code = b.Code "
    sQry = sQry + " Where a.U_CLTCOD = '" & Param01 & "' And b.U_Date = '" & Param02 & "'"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
        Call MDC_Com.MDC_GF_Message("근무일 등록을 하지않습니다.", "E")
        GoTo PH_PY008_MTX01_Exit
    End If
    
    sQry = "EXEC PH_PY008_01 '" & Param01 & "', '" & Param02 & "', '" & Param03 & "', '" & Param04 & "', '" & Param05 & "', '" & Param06 & "', '" & Param07 & "', '" & Param08 & "', '" & oConfirm & "'"
    Call oDS_PH_PY008.ExecuteQuery(sQry)
    
    If (oRecordSet.RecordCount = 0) Then
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
        GoTo PH_PY008_MTX01_Exit
    End If
    
    

    iRow = oForm.DataSources.DataTables.Item(0).Rows.Count
    
    Call PH_PY008_TitleSetting(iRow)
    
'    oDS_PH_PY008.Columns.Item("GetDate").Cells(oRow).Value
    PH_PY008_MTX02 "", 0, ""
    
    If oDS_PH_PY008.Rows.Count > 0 Then
        For i = 0 To oDS_PH_PY008.Rows.Count - 1
            If oDS_PH_PY008.Columns.Item("Chk").Cells(i).VALUE = "Y" Then
                Chkcnt = Chkcnt + 1
            End If
        Next
        oForm.Items("Chkcnt").Specific.VALUE = Chkcnt
    End If
    
    oForm.Update
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Sub
    
PH_PY008_MTX01_Exit:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Sub
PH_PY008_MTX01_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY008_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY008_MTX02(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)

    '//메트릭스 자료를 head에 로드
    
    Dim i       As Long
    Dim sQry    As String
    Dim sRow As Long
    
    Dim Param01 As String
    Dim Param02 As String
    Dim Param03 As String
    
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY008_MTX02_Error

    Call oForm.Freeze(True)
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    sRow = oRow
    
    
    Param01 = oForm.Items("SCLTCOD").Specific.VALUE
    Param02 = oDS_PH_PY008.Columns.Item("GetDate").Cells(oRow).VALUE
    Param03 = oDS_PH_PY008.Columns.Item("MSTCOD").Cells(oRow).VALUE
    
    If Trim(Param02) = "" Then
    
        Set oCombo = oForm.Items("TeamCode").Specific
        oCombo.Select "", psk_ByValue
        
        oForm.Items("MSTCOD").Specific.VALUE = ""
        oForm.Items("FullName").Specific.VALUE = ""
        
        oForm.DataSources.UserDataSources.Item("GetTime").VALUE = "0000"
        oForm.DataSources.UserDataSources.Item("OffTime").VALUE = "0000"
    
        Call PH_PY008_Time_ReSet
        
        oForm.Items.Item("PosDate").Enabled = True
        oForm.Items.Item("TeamCode").Enabled = True
        oForm.Items.Item("RspCode").Enabled = True
        oForm.Items.Item("ClsCode").Enabled = True
        oForm.Items.Item("MSTCOD").Enabled = True
        oForm.Items.Item("FullName").Enabled = True
        
        
        
        
        GoTo PH_PY008_MTX02_Exit
    End If
    sQry = "EXEC PH_PY008_02 '" & Param01 & "', '" & Param02 & "', '" & Param03 & "'"
    Call oRecordSet.DoQuery(sQry)
    
    If (oRecordSet.RecordCount = 0) Then
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
        GoTo PH_PY008_MTX02_Exit
    End If
    
    
    
'    oForm.Items("PosDate").Specific.Value = Format(oRecordset.Fields("PosDate").Value, "YYYYMMDD")
        
    oForm.DataSources.UserDataSources.Item("PosDate").VALUE = Format(oRecordSet.Fields("PosDate").VALUE, "YYYYMMDD")
        
    Set oCombo = oForm.Items("SDayOff").Specific
    oCombo.Select oRecordSet.Fields("DayOff").VALUE, psk_ByValue
        
    oForm.DataSources.UserDataSources.Item("SDay").VALUE = oRecordSet.Fields("Day").VALUE
    oForm.DataSources.UserDataSources.Item("MSTCOD").VALUE = oRecordSet.Fields("MSTCOD").VALUE
    oForm.DataSources.UserDataSources.Item("FullName").VALUE = oRecordSet.Fields("FullName").VALUE
        
'    oForm.Items("SDay").Specific.Value = oRecordset.Fields("Day").Value
'    oForm.Items("MSTCOD").Specific.Value = oRecordset.Fields("MSTCOD").Value
'    oForm.Items("FullName").Specific.Value = oRecordset.Fields("FullName").Value
        
'    '//부서
    Set oCombo = oForm.Items("TeamCode").Specific
    
    If oCombo.ValidValues.Count > 0 Then
        For i = oCombo.ValidValues.Count - 1 To 0 Step -1
            oCombo.ValidValues.Remove i, psk_Index
        Next i
        oCombo.ValidValues.Add "", ""
    End If
    
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
    sQry = sQry & " WHERE Code = '1' "
    sQry = sQry & " Order By U_Seq"
    Call SetReDataCombo(oForm, sQry, oCombo, "Y")
    
    oCombo.Select oRecordSet.Fields("TeamCode").VALUE, psk_ByValue
'
'    '//담당
    Set oCombo = oForm.Items("RspCode").Specific
    
    If oCombo.ValidValues.Count > 0 Then
        For i = oCombo.ValidValues.Count - 1 To 0 Step -1
            oCombo.ValidValues.Remove i, psk_Index
        Next i
        oCombo.ValidValues.Add "", ""
    End If
    
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
    sQry = sQry & " WHERE Code = '2' "
    sQry = sQry & " Order By U_Seq"
    Call SetReDataCombo(oForm, sQry, oCombo, "Y")
    
    oCombo.Select oRecordSet.Fields("RspCode").VALUE, psk_ByValue
    
'    '//반
    Set oCombo = oForm.Items("ClsCode").Specific
    
    If oCombo.ValidValues.Count > 0 Then
        For i = oCombo.ValidValues.Count - 1 To 0 Step -1
            oCombo.ValidValues.Remove i, psk_Index
        Next i
        oCombo.ValidValues.Add "", ""
    End If
    
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
    sQry = sQry & " WHERE Code = '9' "
    sQry = sQry & " Order By U_Seq"
    Call SetReDataCombo(oForm, sQry, oCombo, "Y")
    
    oCombo.Select oRecordSet.Fields("ClsCode").VALUE, psk_ByValue
    
    '//근무형태
    Set oCombo = oForm.Items("ShiftDat").Specific
    oCombo.Select oRecordSet.Fields("ShiftDat").VALUE, psk_ByValue
    
    '//근무조
    Set oCombo = oForm.Items("GNMUJO").Specific
    If oCombo.ValidValues.Count > 0 Then
        For i = oCombo.ValidValues.Count - 1 To 0 Step -1
            oCombo.ValidValues.Remove i, psk_Index
        Next i
        oCombo.ValidValues.Add "", ""
    End If
    
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
    sQry = sQry & " WHERE Code = 'P155' AND U_Char1 = '" & oForm.Items.Item("ShiftDat").Specific.VALUE & "'"
    sQry = sQry & " And U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY U_Code"
    
'    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P155' AND U_UseYN= 'Y'"
    Call SetReDataCombo(oForm, sQry, oCombo, "Y")
    
    oCombo.Select oRecordSet.Fields("GNMUJO").VALUE, psk_ByValue
    
    '//근태구분
    Set oCombo = oForm.Items("WorkType").Specific
    oCombo.Select oRecordSet.Fields("WorkType").VALUE, psk_ByValue
    
    
    
    
    oForm.DataSources.UserDataSources.Item("GetDate").VALUE = Format(oRecordSet.Fields("GetDate").VALUE, "YYYYMMDD")
'    oForm.Items("GetDate").Specific.Value = Format(oRecordset.Fields("GetDate").Value, "YYYYMMDD")
'    oForm.Items("GetTime").Specific.Value = Right("0" + CStr(oRecordset.Fields("GetTime").Value), 4)
    oForm.DataSources.UserDataSources.Item("GetTime").VALUE = Right("0" + CStr(oRecordSet.Fields("GetTime").VALUE), 4)
    oForm.DataSources.UserDataSources.Item("OffDate").VALUE = Format(oRecordSet.Fields("OffDate").VALUE, "YYYYMMDD")
    oForm.DataSources.UserDataSources.Item("OffTime").VALUE = Right("0" + CStr(oRecordSet.Fields("OffTime").VALUE), 4)
'    oForm.Items("OffDate").Specific.Value = Format(oRecordset.Fields("OffDate").Value, "YYYYMMDD")
'    oForm.Items("OffTime").Specific.Value = Right("0" + CStr(oRecordset.Fields("OffTime").Value), 4)
    
    '//부서
'    Set oCombo = oForm.Items("TeamCode").Specific
'    oCombo.Select oRecordset.Fields("TeamCode").Value, psk_ByValue
'
'    '//담당
'    Set oCombo = oForm.Items("RspCode").Specific
'    oCombo.Select oRecordset.Fields("RspCode").Value, psk_ByValue
    
    oForm.DataSources.UserDataSources.Item("Base").VALUE = oRecordSet.Fields("Base").VALUE
    oForm.DataSources.UserDataSources.Item("Extend").VALUE = oRecordSet.Fields("Extend").VALUE
    
    oForm.DataSources.UserDataSources.Item("Special").VALUE = oRecordSet.Fields("Special").VALUE
    oForm.DataSources.UserDataSources.Item("SpExtend").VALUE = oRecordSet.Fields("SpExtend").VALUE
    oForm.DataSources.UserDataSources.Item("Midnight").VALUE = oRecordSet.Fields("Midnight").VALUE
    oForm.DataSources.UserDataSources.Item("LateTo").VALUE = oRecordSet.Fields("LateTo").VALUE
    
    oForm.DataSources.UserDataSources.Item("EarlyOff").VALUE = oRecordSet.Fields("EarlyOff").VALUE
    oForm.DataSources.UserDataSources.Item("EarlyTo").VALUE = oRecordSet.Fields("EarlyTo").VALUE
    
    oForm.DataSources.UserDataSources.Item("SEarlyTo").VALUE = oRecordSet.Fields("SEarlyTo").VALUE '휴일조출(2013.06.17 송명규 추가)
    
    oForm.DataSources.UserDataSources.Item("EducTran").VALUE = oRecordSet.Fields("EducTran").VALUE
    oForm.DataSources.UserDataSources.Item("GoOut").VALUE = oRecordSet.Fields("GoOut").VALUE
    
    oForm.DataSources.UserDataSources.Item("ActCode").VALUE = oRecordSet.Fields("ActCode").VALUE
    oForm.DataSources.UserDataSources.Item("ActText").VALUE = oRecordSet.Fields("ActText").VALUE
    
    Set oCombo = oForm.Items("DangerCD").Specific
    oCombo.Select oRecordSet.Fields("DangerCD").VALUE, psk_ByValue
    
'    oForm.DataSources.UserDataSources.Item("DangerNu").Value = oRecordset.Fields("DangerNu").Value
    oForm.DataSources.UserDataSources.Item("Rotation").VALUE = oRecordSet.Fields("Rotation").VALUE
    

    oForm.DataSources.UserDataSources.Item("RToDate").VALUE = oRecordSet.Fields("RToDate").VALUE
    oForm.DataSources.UserDataSources.Item("RToTime").VALUE = oRecordSet.Fields("RToTime").VALUE
    oForm.DataSources.UserDataSources.Item("ROffDate").VALUE = oRecordSet.Fields("ROffDate").VALUE
    oForm.DataSources.UserDataSources.Item("ROffTime").VALUE = oRecordSet.Fields("ROffTime").VALUE
    
    
    oForm.DataSources.UserDataSources.Item("ROutTime").VALUE = oRecordSet.Fields("ROutTime").VALUE
    oForm.DataSources.UserDataSources.Item("RInTime").VALUE = oRecordSet.Fields("RInTime").VALUE
    
'    oForm.DataSources.UserDataSources.Item("Confirm").Value = oRecordSet.Fields("Confirm").Value
    '//근태구분
    Set oCombo = oForm.Items("Confirm").Specific
    oCombo.Select oRecordSet.Fields("Confirm").VALUE, psk_ByValue
    
    Set oCombo = oForm.Items("RotateYN").Specific
    oCombo.Select oRecordSet.Fields("RotateYN").VALUE, psk_ByValue
'    oForm.DataSources.UserDataSources.Item("RotateYN").VALUE = oRecordSet.Fields("RotateYN").VALUE
    
    '비고(2013.11.19 송명규 추가)
    oForm.DataSources.UserDataSources.Item("Comment").VALUE = oRecordSet.Fields("Comment").VALUE
    
    oForm.Items.Item("PosDate").Enabled = False
    oForm.Items.Item("TeamCode").Enabled = False
    oForm.Items.Item("RspCode").Enabled = False
    oForm.Items.Item("ClsCode").Enabled = False
    oForm.Items.Item("MSTCOD").Enabled = False
    oForm.Items.Item("FullName").Enabled = False
    
    oForm.Update
    
    
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Sub
PH_PY008_MTX02_Exit:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    
    Exit Sub
PH_PY008_MTX02_Error:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY008_MTX02_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY008_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY008_Validate_Error
    PH_PY008_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY008A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", bmt_Short, True
        PH_PY008_Validate = False
        GoTo PH_PY008_Validate_Exit
    End If
'
    If ValidateType = "수정" Then

    ElseIf ValidateType = "행삭제" Then

    ElseIf ValidateType = "취소" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY008_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY008_Validate_Error:
    PH_PY008_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY008_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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

Private Function PH_PY008_DaySelect(ByVal Day As String) As String

    '//요일 Select
    
    Dim sQry    As String
    
   
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY008_DaySelect_Error

    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    sQry = "SELECT datepart(weekday, '" & Day & "')"

    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
       Select Case oRecordSet.Fields(0).VALUE
            Case "1"
                Day = "일"
            Case "2"
                Day = "월"
            Case "3"
                Day = "화"
            Case "4"
                Day = "수"
            Case "5"
                Day = "목"
            Case "6"
                Day = "금"
            Case "7"
                Day = "토"
       End Select
       
    End If
    
    If (oRecordSet.RecordCount = 0) Then
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
        GoTo PH_PY008_DaySelect_Exit
    Else
        PH_PY008_DaySelect = Day
    End If
    
    Set oRecordSet = Nothing
    Exit Function
PH_PY008_DaySelect_Exit:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
PH_PY008_DaySelect_Error:
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "PH_PY008_DaySelect_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Private Function PH_PY008_DayOffSelect(ByVal Day As String) As String
On Error GoTo PH_PY008_DayOffSelect_Error
    '//평일휴일구분 Select
    
    Dim sQry    As String
    Dim CLTCOD As String '사업장
    
    CLTCOD = oForm.Items("SCLTCOD").Specific.VALUE
       
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    
    sQry = "           SELECT    T1.U_DayType "
    sQry = sQry & " FROM      [@PH_PY003A] AS T0"
    sQry = sQry & "              INNER JOIN"
    sQry = sQry & "              [@PH_PY003B] AS T1"
    sQry = sQry & "                 ON T0.Code = T1.Code"
    sQry = sQry & " WHERE    T1.U_Date = '" & Day & "'"
    sQry = sQry & "              AND T0.U_CLTCOD = '" & CLTCOD & "'"
    
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
       PH_PY008_DayOffSelect = oRecordSet.Fields(0).VALUE
    Else
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
        GoTo PH_PY008_DayOffSelect_Exit
    End If
    
    Set oRecordSet = Nothing
    Exit Function
PH_PY008_DayOffSelect_Exit:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
PH_PY008_DayOffSelect_Error:
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "PH_PY008_DayOffSelect_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY008_SAVE()

    '//데이타 저장
    Dim oCombo      As SAPbouiCOM.ComboBox
    
    Dim i       As Long
    Dim sQry    As String
    Dim ErrNum As Integer
    
    Dim CLTCOD$, DayOff$, Day$, MSTCOD$, FullName$, TeamCode$, RspCode$, Inform05$, ClsCode$, ShiftDat$, GNMUJO$, WorkType$
    Dim GetDate As String
    Dim GetTime As Integer
    Dim OffDate As String
    Dim OffTime As Integer
    Dim Base As Double
    Dim Extend As Double
    Dim Special As Double
    Dim SpExtend As Double
    Dim Midnight As Double
    Dim LateTo As Double
    Dim EarlyOff As Double
    Dim EarlyTo As Double
    Dim SEarlyTo As Double
    Dim EducTran As Double
    Dim GoOut As Double
    Dim GOFrTime As Integer
    Dim GOToTime As Integer
    Dim GOFrTim2 As Integer
    Dim GoToTim2 As Integer
    Dim ActCode$, ActText$, DangerCD$
    Dim DangerNu As Double
    Dim Rotation As Double
    Dim RToDate As String
    Dim RToTime As Integer
    Dim ROffDate As String
    Dim ROffTime As Integer
    Dim ROutTime As Integer
    Dim RInTime As Integer
    Dim Abnomal$, Confirm$, Attend$
    Dim PosDate As String
    
    Dim ShiftDat001A As String '//인사기본 교대조
    Dim bShiftDat As String '//지난주 교대조
    Dim RotateYN As String
    
    Dim CheckData As String
    
    Dim OverTime As Double
    Dim OverTimeYN As String
    
    
    Dim SpecialTot As Double  '//사무기술직 특근24시간 제한
    
    Dim JIGTYP As String '//직원구분02 : 사무기술직
    Dim JIGCOD As String '직급(2014.07.23 송명규 추가)
    Dim PAYTYP As String '//급여형태 2:월급제
    
    Dim Comment As String '비고(2013.11.19 송명규 추가)
    
    Dim Param01 As String
    Dim Param02 As String
    Dim Param03 As String
    Dim Param04 As String
    Dim Param05 As String
    Dim Param06 As String
    Dim Param07 As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY008_SAVE_Error

    Call oForm.Freeze(True)
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    RotateYN = "N" '//야간교대조 인정 초기값 N
    
    MSTCOD = oForm.Items("MSTCOD").Specific.VALUE
    CLTCOD = oForm.Items("SCLTCOD").Specific.VALUE
    PosDate = oForm.Items("PosDate").Specific.VALUE
    DayOff = oForm.Items("SDayOff").Specific.VALUE
    Day = oForm.Items("SDay").Specific.VALUE
    ShiftDat = oForm.Items("ShiftDat").Specific.VALUE
    GNMUJO = oForm.Items("GNMUJO").Specific.VALUE
    WorkType = oForm.Items("WorkType").Specific.VALUE
    GetDate = oForm.Items("GetDate").Specific.VALUE
    GetTime = oForm.Items("GetTime").Specific.VALUE
    OffDate = oForm.Items("OffDate").Specific.VALUE
    OffTime = oForm.Items("OffTime").Specific.VALUE
    Base = Trim(oForm.Items("Base").Specific.VALUE)
    Extend = Trim(oForm.Items("Extend").Specific.VALUE)
    Special = Trim(oForm.Items("Special").Specific.VALUE)
    SpExtend = oForm.Items("SpExtend").Specific.VALUE
    Midnight = oForm.Items("Midnight").Specific.VALUE
    LateTo = oForm.Items("LateTo").Specific.VALUE
    EarlyOff = oForm.Items("EarlyOff").Specific.VALUE
    EarlyTo = oForm.Items("EarlyTo").Specific.VALUE
    SEarlyTo = oForm.Items("SEarlyTo").Specific.VALUE
    EducTran = oForm.Items("EducTran").Specific.VALUE
    GoOut = oForm.Items("GoOut").Specific.VALUE
    GOFrTime = oForm.Items("GOFrTime").Specific.VALUE
    GOToTime = oForm.Items("GOToTime").Specific.VALUE
    GOFrTim2 = oForm.Items("GOFrTim2").Specific.VALUE
    GoToTim2 = oForm.Items("GOToTim2").Specific.VALUE
    
    RToDate = oForm.Items("RToDate").Specific.VALUE
    RToTime = oForm.Items("RToTime").Specific.VALUE
    ROffDate = oForm.Items("ROffDate").Specific.VALUE
    ROffTime = oForm.Items("ROffTime").Specific.VALUE
    ROutTime = oForm.Items("ROutTime").Specific.VALUE
    RInTime = oForm.Items("RInTime").Specific.VALUE
'    ActCode = oForm.Items("ActCode").Specific.Value
'    ActText = oForm.Items("ActText").Specific.Value
'    DangerCD = oForm.Items("DangerCD").Specific.Value
'    Rotation = oForm.Items("Rotation").Specific.Value
    Confirm = oForm.Items("Confirm").Specific.VALUE
    Comment = oForm.Items("Comment").Specific.VALUE
    
    
    
    If Trim(Confirm) = "" Then
        Confirm = "N"
    End If
    
    If GNMUJO = "" Then
        Call MDC_Com.MDC_GF_Message("근무조를 입력하지 않습니다.", "E")
        GoTo PH_PY008_SAVE_Exit
    End If
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet.RecordCount, False)
    
    
    If oForm.DataSources.UserDataSources("Chk").VALUE = "N" Then
        '//전체적용
        If oDS_PH_PY008.Rows.Count > 0 Then
            For i = 0 To oDS_PH_PY008.Rows.Count - 1
                If oDS_PH_PY008.Columns.Item("Chk").Cells(i).VALUE = "Y" Then
                    
                    
                    CheckData = Trim(oDS_PH_PY008.Columns.Item("Day").Cells(i).VALUE)
                    MSTCOD = Trim(oDS_PH_PY008.Columns.Item("MSTCOD").Cells(i).VALUE)
                    FullName = Trim(oDS_PH_PY008.Columns.Item("FullName").Cells(i).VALUE)
                    
                    '//지난주 교대조
                    bShiftDat = Trim(oDS_PH_PY008.Columns.Item("bShiftDat").Cells(i).VALUE)
                    
                    '//사원의 직무와 위해 조회
                    sQry = "select Activity = a.U_Activity, "
                    sQry = sQry + " ActText = b.CodeNm, "
                    sQry = sQry + " DangerCD = b.DangerCD, "
                    sQry = sQry + " ShiftDat = a.U_ShiftDat, "
                    sQry = sQry + " JIGTYP = a.U_JIGTYP, "
                    sQry = sQry + " JIGCOD = a.U_JIGCOD, " '직급(2014.07.23 송명규 추가)
                    sQry = sQry + " PAYTYP = a.U_PAYTYP, "
                    sQry = sQry + " TeamCode = a.U_TeamCode, "
                    sQry = sQry + " RspCode = a.U_RspCode, "
                    sQry = sQry + " ClsCode = a.U_ClsCode, "
                    sQry = sQry + " inform05 = a.u_inform05 "
                    sQry = sQry + " FROM [@PH_PY001A] a Left Join (Select Code = t1.U_Code, CodeNm = t1.U_CodeNm, DangerCD = t1.U_Char1 "
                    sQry = sQry + " FROM [@PS_HR200H] t Inner Join [@PS_HR200L] t1 On t.Code = t1.Code "
                    sQry = sQry + " Where t1.Code = 'P127' and U_Char2 = '" & CLTCOD & "' ) b On a.U_Activity = b.Code"
                    sQry = sQry + " Where a.Code = '" & MSTCOD & "'"
                    oRecordSet.DoQuery sQry
                        
                    ActCode = oRecordSet.Fields("Activity").VALUE
                    ActText = oRecordSet.Fields("ActText").VALUE
                    DangerCD = oRecordSet.Fields("DangerCD").VALUE
                        
                    JIGTYP = oRecordSet.Fields("JIGTYP").VALUE '//직원구분
                    JIGCOD = oRecordSet.Fields("JIGCOD").VALUE '직급
                    PAYTYP = oRecordSet.Fields("PAYTYP").VALUE '//급여구분
                    
                    TeamCode = oRecordSet.Fields("TeamCode").VALUE '//팀
                    RspCode = oRecordSet.Fields("RspCode").VALUE '//담당
                    ClsCode = oRecordSet.Fields("ClsCode").VALUE '//반
                    Inform05 = oRecordSet.Fields("inform05").VALUE '//반
                    ShiftDat001A = oRecordSet.Fields("ShiftDat").VALUE '//인사기본 근무조
                    
                    
                    '/////// 1교대자가 2교대야간 근무를 근무했을시 1주일 후 같은 요일에 야간근무인정
                    '/////// 하루 야간근무시 2일 야근교대수당을 지급 2016/01/28 NGY추가
                    If ShiftDat001A = "1" Then
                        If bShiftDat = "2" Then
                            RotateYN = "Y"
                        Else
                            RotateYN = "N"
                        End If
                    End If
                    '//////////////////////////////////////////////////////////////////////////////
                    
                    If CheckData = "" Then
                        '//근태구분이 없으면 전체입력대상
                        
                        Abnomal = "Y" '//근태이상자
                        
                        
                        If oForm.Items("SDay").Specific.VALUE = "월" Then
                            '//월요일은 제외
                            OverTime = 0
                        Else
                            OverTime = PH_PY008_OverTimeCheck(PosDate, MSTCOD)
                        End If
                        ' 스페셜 추가 OVER TIME 주 BASE SPECIAL 도 추가함 (주52)
                        If Base + Special + Extend + SpExtend + EarlyTo + SEarlyTo + OverTime > 52 Then
                        
                            If Inform05 = "Y" Then
                            
                                If Sbo_Application.MessageBox(" 주 52시간 연장초과 했습니다. 그래도 입력하시겠습니까?", "2", "예", "아니오") = "1" Then
                                    OverTime = 0
                                End If
                            
                            Else
                            
                                ErrNum = 1
                                GoTo PH_PY008_SAVE_Error
                                
                            End If
'                            Sbo_Application.SetStatusBarMessage "주 12시간 연장초과 했습니다. 입력할 수 없습니다.", bmt_Short, True
'                            If Sbo_Application.MessageBox(" 주 12시간 연장초과 했습니다. 그래도 입력하시겠습니까? ?", "2", "예", "아니오") = "1" Then
'                                OverTime = 0
'                            End If
                        End If
                        
                        If Base + Special + Extend + SpExtend + EarlyTo + SEarlyTo + OverTime <= 52 Then
                            '//지각, 조퇴, 외출시간이 있으면 교대일수 없다.
                            If LateTo + EarlyOff + GoOut > 0.1 Then
                                Rotation = 0
                            Else
                                Rotation = 1
                            End If
                            
                            If Left(WorkType, 1) = "F" Or WorkType = "D11" Then  '//휴직, 무급휴가
                                Rotation = 0
                            End If
                            
                            
                            If (JIGTYP = "02" Or JIGTYP = "03" Or JIGCOD = "73") And PAYTYP = "2" Then '//사무기술직(사무계약직(73) 추가, 2014.07.23 송명규) 월급제 연장근무시간 자동계산 월 20시간
                                If CLTCOD <> "2" Then '주52시간 근무 시행(2018.07.01 부)으로 월급제 기능 제한(2018.07.05 송명규), 일단 부산사업장만 시행
                                    Extend = PH_PY008_Extend15(MSTCOD)
                                    
                                    SpecialTot = PH_PY008_Special(MSTCOD)
                                    
                                    If SpecialTot >= 24 Then
                                        Special = 0
                                    Else
                                        If (Special + SpecialTot) >= 24 Then
                                            Special = Special - (Special + SpecialTot - 24)
                                        End If
                                    End If
                                Else
                                    Extend = Trim(oForm.Items("Extend").Specific.VALUE)
                                End If
                            Else
                                Extend = Trim(oForm.Items("Extend").Specific.VALUE)
                            End If
                            
                            '무단결근, 유계결근, 휴직, 무급휴가는 위해 수당 없다.
                            If WorkType = "A01" Or WorkType = "A02" Or Left(WorkType, 1) = "F" Or WorkType = "D11" Then
                                DangerCD = ""
                            
                            Else
                                '//전문직, 계약직 이며 연봉제가 아니고 위해코드가 없으면 위해코드를 기타로..
                                If (JIGTYP = "04" Or JIGTYP = "05") And PAYTYP <> "1" And JIGCOD <> "73" And DangerCD = "" Then
                                    If CLTCOD = "1" Then '창원사업장만 적용(2013.09.30 송명규 추가)
                                        DangerCD = "31"
                                    End If
                                End If
                            End If
                            
                            '근무시간이 있으나 4시간미만이면 위해수당 없다.
                            If Base + Special > 0 Then
                                If Special + SpExtend + Base + Extend < 4 Then
                                    DangerCD = ""
                                End If
                            Else
                                '//근무시간이 없으면 위해코드를 기타로.
                                If CLTCOD = "1" Then '창원사업장만 적용(2015.04.24 노근용 추가)
                                    If Trim(DangerCD) <> "" Then
                                        
                                        DangerCD = "31"
                                    End If
                                End If
                            End If
                            
                            If DangerCD = "" Then
                                DangerNu = 0
                            Else
                                DangerNu = 1
                            End If
                            
                            If CLTCOD = "3" Then  '//포장사업팀은 위해수당 자동계산 안함
                                DangerCD = ""
                                DangerNu = 0
                            End If
                            
                            sQry = "INSERT INTO ZPH_PY008"
                            sQry = sQry & " ("
                            sQry = sQry & " CLTCOD,"
                            sQry = sQry & " PosDate,"
                            sQry = sQry & " DayOff,"
                            sQry = sQry & " Day,"
                            sQry = sQry & " MSTCOD,"
                            sQry = sQry & " FullName,"
                            sQry = sQry & " TeamCode,"
                            sQry = sQry & " RspCode,"
                            sQry = sQry & " ClsCode,"
                            sQry = sQry & " ShiftDat,"
                            sQry = sQry & " GNMUJO,"
                            sQry = sQry & " WorkType,"
                            sQry = sQry & " GetDate,"
                            sQry = sQry & " GetTime,"
                            sQry = sQry & " OffDate,"
                            sQry = sQry & " OffTime,"
                            sQry = sQry & " Base,"
                            sQry = sQry & " Extend,"
                            sQry = sQry & " Special,"
                            sQry = sQry & " SpExtend,"
                            sQry = sQry & " Midnight,"
                            sQry = sQry & " LateTo,"
                            sQry = sQry & " EarlyOff,"
                            sQry = sQry & " EarlyTo,"
                            sQry = sQry & " SEarlyTo,"
                            sQry = sQry & " EducTran,"
                            sQry = sQry & " GoOut,"
                            sQry = sQry & " GOFrTime,"
                            sQry = sQry & " GOToTime,"
                            sQry = sQry & " GOFrTim2,"
                            sQry = sQry & " GOToTim2,"
                            sQry = sQry & " ActCode,"
                            sQry = sQry & " ActText,"
                            sQry = sQry & " DangerCD,"
                            sQry = sQry & " DangerNu,"
                            sQry = sQry & " Rotation,"
                            sQry = sQry & " RToDate,"
                            sQry = sQry & " RToTime,"
                            sQry = sQry & " ROffDate,"
                            sQry = sQry & " ROffTime,"
                            sQry = sQry & " ROutTime,"
                            sQry = sQry & " RInTime,"
                            sQry = sQry & " Abnomal,"
                            sQry = sQry & " Attend,"
                            sQry = sQry & " Confirm,"
                            sQry = sQry & " Comment,"
                            sQry = sQry & " RotateYN"
                            sQry = sQry & " ) "
                            sQry = sQry & "VALUES("
                            sQry = sQry & "'" & CLTCOD & "',"
                            sQry = sQry & "'" & PosDate & "',"
                            sQry = sQry & "'" & DayOff & "',"
                            sQry = sQry & "'" & Day & "',"
                            sQry = sQry & "'" & MSTCOD & "',"
                            sQry = sQry & "'" & FullName & "',"
                            sQry = sQry & "'" & TeamCode & "',"
                            sQry = sQry & "'" & RspCode & "',"
                            sQry = sQry & "'" & ClsCode & "',"
                            sQry = sQry & "'" & ShiftDat & "',"
                            sQry = sQry & "'" & GNMUJO & "',"
                            sQry = sQry & "'" & WorkType & "',"
                            sQry = sQry & "'" & GetDate & "',"
                            sQry = sQry & GetTime & ","
                            sQry = sQry & "'" & OffDate & "',"
                            sQry = sQry & OffTime & ","
                            sQry = sQry & Base & ","
                            sQry = sQry & Extend & ","
                            sQry = sQry & Special & ","
                            sQry = sQry & SpExtend & ","
                            sQry = sQry & Midnight & ","
                            sQry = sQry & LateTo & ","
                            sQry = sQry & EarlyOff & ","
                            sQry = sQry & EarlyTo & ","
                            sQry = sQry & SEarlyTo & ","
                            sQry = sQry & EducTran & ","
                            sQry = sQry & GoOut & ","
                            sQry = sQry & GOFrTime & ","
                            sQry = sQry & GOToTime & ","
                            sQry = sQry & GOFrTim2 & ","
                            sQry = sQry & GoToTim2 & ","
                            sQry = sQry & "'" & ActCode & "',"
                            sQry = sQry & "'" & ActText & "',"
                            sQry = sQry & "'" & DangerCD & "',"
                            sQry = sQry & DangerNu & ","
                            sQry = sQry & Rotation & ","
                            sQry = sQry & "'" & RToDate & "',"
                            sQry = sQry & RToTime & ","
                            sQry = sQry & "'" & ROffDate & "',"
                            sQry = sQry & ROffTime & ","
                            sQry = sQry & ROutTime & ","
                            sQry = sQry & RInTime & ","
                            sQry = sQry & "'" & Abnomal & "',"
                            sQry = sQry & "'" & Attend & "',"
                            sQry = sQry & "'" & Confirm & "',"
                            sQry = sQry & "'" & Comment & "',"
                            sQry = sQry & "'" & RotateYN & "'"
                            sQry = sQry & ")"
                            oRecordSet.DoQuery sQry
                        End If
                    Else
                        If oForm.Items("SDay").Specific.VALUE = "월" Then
                            '//월요일은 제외
                            sQry = "SELECT Isnull(U_Inform05,'N') From [@PH_PY001A] Where Code = '" & MSTCOD & "'"
                            oRecordSet.DoQuery sQry

                            OverTimeYN = oRecordSet.Fields(0).VALUE
                            
                            If OverTimeYN = "Y" Then
                                OverTime = -100
                            Else
                                OverTime = 0
                            End If
                        Else
                            OverTime = PH_PY008_OverTimeCheck(PosDate, MSTCOD)
                        End If
                        ' 스페셜 추가 OVER TIME 주 BASE SPECIAL 도 추가함 (주52)
                        If Base + Special + Extend + SpExtend + EarlyTo + SEarlyTo + OverTime > 52 Then
                        
                            If Inform05 = "Y" Then
                            
                                If Sbo_Application.MessageBox(" 주 52시간 연장초과 했습니다. 그래도 입력하시겠습니까? ?", "2", "예", "아니오") = "1" Then
                                    OverTime = 0
                                End If
                            
                            Else
                            
                                ErrNum = 1
                                GoTo PH_PY008_SAVE_Error
                                
                            End If
                            
'                            Sbo_Application.SetStatusBarMessage "주 12시간 연장초과 했습니다. 입력할 수 없습니다.", bmt_Short, True
'                            If Sbo_Application.MessageBox(" 주 12시간 연장초과 했습니다. 그래도 입력하시겠습니까? ?", "2", "예", "아니오") = "1" Then
'                                OverTime = -100
'                            End If
                        End If
                        ' 스페셜 추가 OVER TIME 주 BASE SPECIAL 도 추가함 (주52)
                        If Base + Special + Extend + SpExtend + EarlyTo + SEarlyTo + OverTime <= 52 Then
                            '//지각, 조퇴, 외출시간이 있으면 교대일수 없다.
                            If LateTo + EarlyOff + GoOut > 0.1 Then
                                Rotation = 0
                            Else
                                Rotation = 1
                            End If
                                        
                            If Left(WorkType, 1) = "F" Or WorkType = "D11" Then '//휴직, 무급휴가
                                Rotation = 0
                            End If
                                        
                            If (JIGTYP = "02" Or JIGTYP = "03" Or JIGCOD = "73") And PAYTYP = "2" Then '//사무기술직(사무계약직(73) 추가, 2014.07.23 송명규) 월급제 연장근무시간 자동계산 월 20시간 'JIGTYP = "02" And PAYTYP = "2" Then '//사무기술직 월급제 연장근무시간 자동계산 월 15시간
                                Extend = PH_PY008_Extend15(MSTCOD)
                                
                                SpecialTot = PH_PY008_Special(MSTCOD) '총특근시간
                                
                                If SpecialTot >= 24 Then
                                    Special = 0
                                Else
                                    If (Special + SpecialTot) >= 24 Then
                                        Special = Special - (Special + SpecialTot - 24)
                                    End If
                                End If
                                
                            Else
                                Extend = Trim(oForm.Items("Extend").Specific.VALUE)
                            End If
                            
                            '무단결근, 유계결근, 휴직, 무급휴가는 위해 수당 없다.
                            If WorkType = "A01" Or WorkType = "A02" Or Left(WorkType, 1) = "F" Or WorkType = "D11" Then
                                DangerCD = ""
                            
                            Else
                                '//전문직, 계약직 이며 위해코드가 없으면 위해코드를 기타로..
                                ' JIGCOD 73 추가 20180201 황영수 사무계약직의 경우 추가 안되도록 수정 김택근과장 요청
                                If (JIGTYP = "04" Or JIGTYP = "05") And DangerCD = "" And JIGCOD <> "73" Then
                                    If CLTCOD = "1" Then '창원사업장만 적용(2013.10.02 송명규 수정)
                                        DangerCD = "31"
                                    End If
                                End If
                            End If
                            
                            '근무시간이 있으나 4시간미만이면 위해수당 없다.
                            If Base + Special > 0 Then
                                If Special + SpExtend + Base + Extend < 4 Then
                                    DangerCD = ""
                                End If
                            Else
                                '//근무시간이 없으면 위해코드를 기타로.
                                 ' JIGCOD 73 추가 20180201 황영수 사무계약직의 경우 추가 안되도록 수정 김택근과장 요청
                                If CLTCOD = "1" And JIGCOD <> "73" Then '창원사업장만 적용(2015.04.24 노근용 추가)
                                    If Trim(DangerCD) <> "" Then
                                        DangerCD = "31"
                                    End If
                                End If
                            End If
                            
'                            If CLTCOD = "1" Then
'                                '창원사업장의 위해코드 무단결근, 유계결근외
'                                If WorkType = "A01" Or WorkType = "A02" Then
'                                    DangerCD = ""
'                                End If
'                            Else
'                                If Left(WorkType, 1) <> "H" And Left(WorkType, 1) <> "C" And Special + SpExtend + Base + Extend >= 4 Then
'                                Else
'                                    DangerCD = ""
'                                    '//출장, 훈련
'                                End If
'                            End If
                            
                            If DangerCD = "" Then
                                DangerNu = 0
                            Else
                                DangerNu = 1
                            End If
                            
                            
                            sQry = "Update ZPH_PY008"
                            sQry = sQry & " Set ShiftDat = '" & ShiftDat & "',"
                            sQry = sQry & " GNMUJO = '" & GNMUJO & "',"
                            sQry = sQry & " WorkType = '" & WorkType & "',"
                            sQry = sQry & " GetDate = '" & GetDate & "',"
                            sQry = sQry & " GetTime = " & GetTime & ","
                            sQry = sQry & " OffDate = '" & OffDate & "',"
                            sQry = sQry & " OffTime = " & OffTime & ","
                            sQry = sQry & " Base = " & Base & ","
                            sQry = sQry & " Extend = " & Extend & ","
                            sQry = sQry & " Special = " & Special & ","
                            sQry = sQry & " SpExtend = " & SpExtend & ","
                            sQry = sQry & " Midnight = " & Midnight & ","
                            sQry = sQry & " LateTo = " & LateTo & ","
                            sQry = sQry & " EarlyOff = " & EarlyOff & ","
                            sQry = sQry & " EarlyTo = " & EarlyTo & ","
                            sQry = sQry & " SEarlyTo = " & SEarlyTo & ","
                            sQry = sQry & " EducTran = " & EducTran & ","
                            sQry = sQry & " GoOut = " & GoOut & ","
                            sQry = sQry & " GOFrTime = " & GOFrTime & ","
                            sQry = sQry & " GOToTime = " & GOToTime & ","
                            sQry = sQry & " GOFrTim2 = " & GOFrTim2 & ","
                            sQry = sQry & " GOToTim2 = " & GoToTim2 & ","
                            sQry = sQry & " ActCode = '" & ActCode & "',"
                            sQry = sQry & " ActText = '" & ActText & "',"
                            sQry = sQry & " DangerCD = '" & DangerCD & "',"
                            sQry = sQry & " DangerNu = " & DangerNu & ","
                            sQry = sQry & " Rotation = " & Rotation & ","
                            sQry = sQry & " RToDate = '" & RToDate & "',"
                            sQry = sQry & " RToTime = " & RToTime & ","
                            sQry = sQry & " ROffDate = '" & ROffDate & "',"
                            sQry = sQry & " ROffTime = " & ROffTime & ","
                            sQry = sQry & " ROutTime = " & ROutTime & ","
                            sQry = sQry & " RInTime = " & RInTime & ","
                            sQry = sQry & " Confirm = '" & Confirm & "',"
                            sQry = sQry & " Comment = '" & Comment & "',"
                            sQry = sQry & " RotateYN = '" & RotateYN & "'"
                            sQry = sQry & " Where CLTCOD = '" & CLTCOD & "'"
                            sQry = sQry & " And PosDate = '" & PosDate & "'"
                            sQry = sQry & " And MSTCOD = '" & MSTCOD & "'"
                            
                            oRecordSet.DoQuery sQry
                        
                        End If
                    End If
 
                End If
                ProgressBar01.VALUE = ProgressBar01.VALUE + 1
                ProgressBar01.Text = ProgressBar01.VALUE & "/" & oDS_PH_PY008.Rows.Count & "건 저장중...!"
            Next i
        End If
    Else
    ' //근태이상자 적용

        '//사원의 직무와 위해 조회
        sQry = "select Activity = a.U_Activity, "
        sQry = sQry + " ActText = b.CodeNm, "
        sQry = sQry + " DangerCD = b.DangerCD, "
        sQry = sQry + " JIGTYP = a.U_JIGTYP, "
        sQry = sQry + " PAYTYP = a.U_PAYTYP, "
        sQry = sQry + " JIGCOD = a.U_JIGCOD "
        sQry = sQry + " FROM [@PH_PY001A] a Left Join (Select Code = t1.U_Code, CodeNm = t1.U_CodeNm, DangerCD = t1.U_Char1 "
        sQry = sQry + " FROM [@PS_HR200H] t Inner Join [@PS_HR200L] t1 On t.Code = t1.Code "
        sQry = sQry + " Where t1.Code = 'P127' and U_Char2 = '" & CLTCOD & "' ) b On a.U_Activity = b.Code"
        sQry = sQry + " Where a.Code = '" & MSTCOD & "'"
        oRecordSet.DoQuery sQry

        ActCode = oRecordSet.Fields(0).VALUE
        ActText = oRecordSet.Fields(1).VALUE
        DangerCD = oRecordSet.Fields(2).VALUE

        JIGTYP = oRecordSet.Fields(3).VALUE '//직원구분
        PAYTYP = oRecordSet.Fields(4).VALUE '//급여구분
        JIGCOD = oRecordSet.Fields(5).VALUE '직급


        If oForm.Items("SDay").Specific.VALUE = "월" Then
            '//월요일은 제외
            OverTime = 0
        Else
            OverTime = PH_PY008_OverTimeCheck(PosDate, MSTCOD)
        End If
        
        If Base + Special + Extend + SpExtend + EarlyTo + SEarlyTo + OverTime >= 52 Then
            If Sbo_Application.MessageBox(" 주 52시간을 초과 했습니다. 그래도 입력하시겠습니까? ?", "2", "예", "아니오") = "1" Then
                OverTime = 0
            End If
        End If
                        
        If Base + Special + Extend + SpExtend + EarlyTo + SEarlyTo + OverTime <= 52 Then
            '//지각, 조퇴, 외출시간이 있으면 교대일수 없다.
            If LateTo + EarlyOff + GoOut > 0.1 Then
                Rotation = 0
            Else
                Rotation = 1
            End If
            
            If Left(WorkType, 1) = "F" Or WorkType = "D11" Then '//휴직
                Rotation = 0
            End If
                            
            If (JIGTYP = "02" Or JIGTYP = "03" Or JIGCOD = "73") And PAYTYP = "2" Then '//사무기술직(사무계약직(73) 추가, 2014.07.23 송명규) 월급제 연장근무시간 자동계산 월 20시간 'JIGTYP = "02" And PAYTYP = "2" Then '//사무기술직 월급제 연장근무시간 자동계산 월 15시간
                Extend = PH_PY008_Extend15(MSTCOD)

            Else
                Extend = Trim(oForm.Items("Extend").Specific.VALUE)
            End If
            
            If CLTCOD = "1" Then
                '무단결근, 유계결근, 휴직, 무급휴가는 위해 수당 없다.
                If WorkType = "A01" Or WorkType = "A02" Or Left(WorkType, 1) = "F" Or WorkType = "D11" Then
                    DangerCD = ""
                End If
            Else
                            
                If Left(WorkType, 1) <> "H" And Left(WorkType, 1) <> "C" And Special + SpExtend + Base + Extend >= 4 Then
                Else
                    DangerCD = ""
                    '//출장, 훈련
                End If
            End If
'
            sQry = "Update ZPH_PY008"
            sQry = sQry & " Set ShiftDat = '" & ShiftDat & "',"
            sQry = sQry & " GNMUJO = '" & GNMUJO & "',"
            sQry = sQry & " WorkType = '" & WorkType & "',"
            sQry = sQry & " GetDate = '" & GetDate & "',"
            sQry = sQry & " GetTime = " & GetTime & ","
            sQry = sQry & " OffDate = '" & OffDate & "',"
            sQry = sQry & " OffTime = " & OffTime & ","
            sQry = sQry & " Base = " & Base & ","
            sQry = sQry & " Extend = " & Extend & ","
            sQry = sQry & " Special = " & Special & ","
            sQry = sQry & " SpExtend = " & SpExtend & ","
            sQry = sQry & " Midnight = " & Midnight & ","
            sQry = sQry & " LateTo = " & LateTo & ","
            sQry = sQry & " EarlyOff = " & EarlyOff & ","
            sQry = sQry & " EarlyTo = " & EarlyTo & ","
            sQry = sQry & " SEarlyTo = " & SEarlyTo & ","
            sQry = sQry & " EducTran = " & EducTran & ","
            sQry = sQry & " GoOut = " & GoOut & ","
            sQry = sQry & " GOFrTime = " & GOFrTime & ","
            sQry = sQry & " GOToTime = " & GOToTime & ","
            sQry = sQry & " GOFrTim2 = " & GOFrTim2 & ","
            sQry = sQry & " GOToTim2 = " & GoToTim2 & ","
            sQry = sQry & " Rotation = " & Rotation & ","
            sQry = sQry & " RToDate = '" & RToDate & "',"
            sQry = sQry & " RToTime = " & RToTime & ","
            sQry = sQry & " ROffDate = '" & ROffDate & "',"
            sQry = sQry & " ROffTime = " & ROffTime & ","
            sQry = sQry & " ROutTime = " & ROutTime & ","
            sQry = sQry & " RInTime = " & RInTime & ","
            sQry = sQry & " Confirm = '" & Confirm & "'"
            sQry = sQry & " Where CLTCOD = '" & CLTCOD & "'"
            sQry = sQry & " And PosDate = '" & PosDate & "'"
            sQry = sQry & " And MSTCOD = '" & MSTCOD & "'"

            oRecordSet.DoQuery sQry
'
        End If
'
    End If
    
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
        Set ProgressBar01 = Nothing
    End If
    
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    
    Set oCombo = oForm.Items("WorkType").Specific
    oCombo.Select oWorkType, psk_ByValue
                    
    If oForm.DataSources.UserDataSources("Chk").VALUE = "N" Then
        Call PH_PY008_MTX01("A")
    Else
        Call PH_PY008_MTX01("S")
    End If

    Exit Sub
    
PH_PY008_SAVE_Exit:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
        Set ProgressBar01 = Nothing
    End If
    Exit Sub
PH_PY008_SAVE_Error:
    Call oForm.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
        Set ProgressBar01 = Nothing
    End If
    
    Set oRecordSet = Nothing
    
    If ErrNum = 1 Then
    
        Sbo_Application.SetStatusBarMessage "주 52시간 연장초과 했습니다. 입력할 수 없습니다. 현재 총 연장시간은 " & Base + Special + Extend + SpExtend + EarlyTo + SEarlyTo + OverTime & "시간 입니다.", bmt_Short, True
    
    Else
    
        Sbo_Application.SetStatusBarMessage "PH_PY008_SAVE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    
    End If
    
End Sub

Private Function PH_PY008_WorkTypeSelect(ByVal Day As String) As String

    '//근태구분 Select
    
    Dim sQry    As String
    
   
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY008_WorkTypeSelect_Error

    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    sQry = "SELECT U_WorkType From [@PH_PY003B] Where U_Date = '" & Day & "'"

    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
       PH_PY008_WorkTypeSelect = oRecordSet.Fields(0).VALUE
    Else
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
        GoTo PH_PY008_WorkTypeSelect_Exit
    End If
    
    Set oRecordSet = Nothing
    Exit Function
PH_PY008_WorkTypeSelect_Exit:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
PH_PY008_WorkTypeSelect_Error:
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "PH_PY008_WorkTypeSelect_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Private Function PH_PY008_Time_Calc_Main(ByVal OffTime As String)

    '//근태구분 Select
    Dim i As Integer
    Dim CLTCOD  As String
    Dim ShiftDat As String '//근무형태
    Dim GNMUJO  As String '//근무조
    Dim DayOff As String '//평일휴일구분
    
    Dim PosDate As String '//기준일
    Dim GetDate As String '//출근일
    Dim OffDate As String '//퇴근일
    Dim GetTime As String '//출근시간
    'Dim OffTime As String '//퇴근시간
    
    Dim FromTime As String
    Dim ToTime As String
    
    Dim hTime1 As Double '//오전10분휴식시간
    Dim hTime2 As Double '//점심휴식시간
    Dim hTime3 As Double '//오후 10분 휴식시간
    Dim hTime4 As Double '//저녁휴식시간
    Dim hTime5 As Double '//야간휴식시간
    
    
    Dim NextDay As String
    Dim TimeType As String
    Dim sQry    As String
    
    Dim STime As String
    Dim ETime As String
    
    Dim EarlyTo As Double
    Dim SEarlyTo As Double
    Dim Base As Double
    Dim Special As Double
    Dim Extend As Double
    Dim SpExtend As Double
    Dim Midnight As Double
    
    Dim WorkType As String
    Dim Team1 As String
    Dim Team2 As String
    Dim Team3 As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY008_Time_Calc_Main_Error

    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    CLTCOD = oForm.Items("SCLTCOD").Specific.VALUE
    ShiftDat = oForm.Items("ShiftDat").Specific.VALUE
    GNMUJO = oForm.Items("GNMUJO").Specific.VALUE
    
    PosDate = oForm.Items("PosDate").Specific.VALUE
    GetDate = oForm.Items("GetDate").Specific.VALUE
    OffDate = oForm.Items("OffDate").Specific.VALUE
    DayOff = oForm.Items("DayOff1").Specific.VALUE
    GetTime = CStr(oForm.Items("GetTime").Specific.VALUE)
    OffTime = CStr(oForm.Items("OffTime").Specific.VALUE)
    
    Team1 = Trim(oForm.Items("Team1").Specific.VALUE)
    Team2 = Trim(oForm.Items("Team2").Specific.VALUE)
    Team3 = Trim(oForm.Items("Team3").Specific.VALUE)
    
    
    GetTime = "0" + GetTime
    GetTime = Right(GetTime, 4)
    
    OffTime = "0" + OffTime
    OffTime = Right(OffTime, 4)
    
    WorkType = oForm.Items("WorkType").Specific.VALUE
    If (GNMUJO = "41") Then
        If (Team1 = "D") Then
             GNMUJO = "21"
        ElseIf (Team1 = "N") Then
             GNMUJO = "22"
        Else
             GNMUJO = ""
        End If
    End If
        
    If (GNMUJO = "42") Then
        If (Team2 = "D") Then
             GNMUJO = "21"
        ElseIf (Team2 = "N") Then
             GNMUJO = "22"
        Else
             GNMUJO = ""
        End If
    End If
            
    If (GNMUJO = "43") Then
        If (Team3 = "D") Then
             GNMUJO = "21"
        ElseIf (Team3 = "N") Then
             GNMUJO = "22"
        Else
             GNMUJO = ""
        End If
    End If
    
    ' 위 근무조를 받아와서 21 OR 22 로 변경함 (주간, 야간 시간표때문 PHPY002) (주52)
    If (ShiftDat = "4") Then
    
        ShiftDat = "2"
        
        If (WorkType = "A00") Then
            DayOff = "1"
        Else
            DayOff = "2"
        End If
        
        sQry = " Select U_TimeType, U_FromTime, U_ToTime, U_NextDay from [@PH_PY002A] a Inner Join [@PH_PY002B] b On a.Code = b.Code "
        sQry = sQry + " Where a.U_CLTCOD = '" & CLTCOD & "'" '//사업부
        sQry = sQry + " and a.U_SType = '" & ShiftDat & "'" '//교대
        sQry = sQry + " and a.U_Shift = '" & GNMUJO & "'"  '//조
        sQry = sQry + " and b.U_DayType = '" & DayOff & "'" '//평일

        ShiftDat = "4"
    Else
        sQry = " Select U_TimeType, U_FromTime, U_ToTime, U_NextDay from [@PH_PY002A] a Inner Join [@PH_PY002B] b On a.Code = b.Code "
        sQry = sQry + " Where a.U_CLTCOD = '" & CLTCOD & "'" '//사업부
        sQry = sQry + " and a.U_SType = '" & ShiftDat & "'" '//교대
        sQry = sQry + " and a.U_Shift = '" & GNMUJO & "'"  '//조
        sQry = sQry + " and b.U_DayType = '" & DayOff & "'" '//평일
    End If
    oRecordSet.DoQuery sQry
    
    If (oRecordSet.RecordCount = 0) Then
'        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
        GoTo PH_PY008_Time_Calc_Main_Exit
    End If
    
    For i = 0 To oRecordSet.RecordCount - 1
        FromTime = CStr(oRecordSet.Fields(1).VALUE)
        FromTime = "0000" + FromTime
        FromTime = Right(FromTime, 4)
        
        ToTime = CStr(oRecordSet.Fields(2).VALUE)
        ToTime = "0000" + ToTime
        ToTime = Right(ToTime, 4)
        
        NextDay = Trim(oRecordSet.Fields(3).VALUE)
        TimeType = Trim(oRecordSet.Fields(0).VALUE)
        
        If NextDay = "" Then
            NextDay = "N"
        End If
        
        If NextDay = "N" Then
            If ToTime = "0000" Then
                ToTime = "2400"
            End If
        End If
        
        Select Case TimeType
            Case "40" '//조출
                If GetDate = PosDate Then '//출근당일이면
                    If GetTime < ToTime Then '//출근시간 < 기준종료시간
                        STime = GetTime
                        ETime = ToTime
                        
                        If DayOff = "1" Then
                            EarlyTo = PH_PY008_Time_Calc(STime, ETime)
                        Else
                            SEarlyTo = PH_PY008_Time_Calc(STime, ETime)
                        End If
                        
                    End If
                End If
            Case "10", "50" '//정상근무시간
                If (GNMUJO = "11" Or GNMUJO = "21" Or GNMUJO = "31" Or GNMUJO = "32") Then
                    Select Case NextDay
                        Case "N" '//당일
                            If PosDate = GetDate Then '//기준일자 = 출근일자
                                '//1교대1조, 2교대 1조당일
                                If GetTime < FromTime Then '//출근시간 < 시작시간
                                    STime = FromTime       '//시작시간
                                Else
                                    STime = GetTime        '//출근시간
                                End If
                                    
                                If GetDate <> OffDate Then '//퇴근일이 틀릴때(다음날 퇴근일때)
                                    ETime = ToTime  '//종료시간
                                Else
                                    If OffTime < ToTime Then '//퇴근시간 < 종료시간
                                        ETime = OffTime      '//퇴근시간
                                    Else
                                        ETime = ToTime       '//종료시간
                                    End If
                                End If
                                
                                If DayOff = 1 Then
                                    Base = Base + PH_PY008_Time_Calc(STime, ETime) '//평일
                                Else
                                    Special = Special + PH_PY008_Time_Calc(STime, ETime) '//휴일
                                End If
                            End If
                        Case "Y"  '//익일
'                            If PosDate = OffDate Then '//기준일자와 퇴근일자가 같으면 계산안함(당일퇴근)
'                                STime = "0000"
'                                ETime = "0000"
'                            Else                      '//익일퇴근일때
'                                If OffTime < ToTime Then  '//퇴근시간 < 종료시간
'                                    STime = "0000"        '//00시 부터
'                                    ETime = OffTime       '//퇴근시간
'                                Else
'                                    STime = "0000"
'                                    ETime = ToTime
'                                End If
'                            End If
'
'                            If DayOff = 1 Then
'                                Base = Base + PH_PY008_Time_Calc(STime, ETime)
'                            Else
'                                Special = Special + PH_PY008_Time_Calc(STime, ETime)
'                            End If
                        End Select
                Else
                    If GNMUJO = "22" Or GNMUJO = "33" Then '//당일 익일 더해야 한다.
                        Select Case NextDay
                            Case "N" '//당일
                                If PosDate = GetDate Then '//기준일 = 출근일
                                    If GetTime < FromTime Then  '//출근시간 < 시작시간
                                        STime = FromTime        '//시작시간
                                    Else
                                        STime = GetTime         '//출근시간
                                    End If
                                    
                                    If GetDate = OffDate Then '//퇴근일이 같으면(기준일 퇴근)
                                        If OffTime < "2400" Then '//퇴근시간 < 24시
                                            ETime = OffTime         '//퇴근시간
                                        Else
                                            ETime = "2400"
                                        End If
                                    Else
                                        ETime = "2400"
                                    End If
                                    
                                    If DayOff = 1 Then
                                        Base = Base + PH_PY008_Time_Calc(STime, ETime)
                                    Else
                                        Special = Special + PH_PY008_Time_Calc(STime, ETime)
                                    End If
                                End If
                                
                            Case "Y" '//익일
                                If PosDate = OffDate Then '//기준일이 같으면 계산안함
                                    STime = "0000"
                                    ETime = "0000"
                                Else
                                    STime = "0000"      '//시작시간은 00시
                                    
                                    If OffTime < ToTime Then '//퇴근시간 < 종료시간
                                        ETime = OffTime        '//퇴근시간
                                    Else
                                        ETime = ToTime          '//종료시간
                                    End If
                                    
                                End If
                                
                                If DayOff = 1 Then
                                    Base = Base + PH_PY008_Time_Calc(STime, ETime)
                                Else
                                    Special = Special + PH_PY008_Time_Calc(STime, ETime)
                                End If
                                    
                        End Select
                        
                    End If
                    
                End If
            Case "65", "66", "15" '//오전, 오후 휴식시간, 점심시간
                If (GNMUJO = "11" Or GNMUJO = "21" Or GNMUJO = "22" Or GNMUJO = "31" Or GNMUJO = "32" Or GNMUJO = "33") Then '//1교대1조 2교대 1조
                    Select Case NextDay
                        Case "N" '//당일
                            If PosDate <> GetDate Then '//당일 출근이 아니면
                                STime = "0000"
                                ETime = "0000"
                            Else
                                If PosDate = OffDate Then '//당일퇴근
                                    If GetTime < FromTime Then  '//출근시간 < 시작시간
                                        If FromTime <= OffTime Then '//시작시간 <= 퇴근시간
                                            STime = FromTime  '//시작시간
                                            If ToTime < OffTime Then    '//종료시간 < 퇴근시간
                                                ETime = ToTime
                                            Else
                                                ETime = OffTime
                                            End If
                                        Else
                                            STime = "0000"
                                            ETime = "0000"
                                        End If
                                    Else
                                        If GetTime > ToTime Then    '//출근시간 > 종료시간
                                            STime = "0000"
                                            ETime = "0000"
                                        Else
                                            STime = GetTime    '//출근시간
                                            If ToTime < OffTime Then    '//종료시간 < 퇴근시간
                                                ETime = ToTime          '//종료시간
                                            Else
                                                ETime = OffTime         '//퇴근시간
                                            End If
                                        End If
                                    End If
                                Else '//다음날 퇴근
                                    If ToTime < GetTime Then '//종료시간 < 출근시간
                                        STime = "0000"
                                        ETime = "0000"
                                    Else
                                        If GetTime < FromTime Then   '//출근시간 < 시작시간
                                            STime = FromTime
                                            ETime = ToTime
                                            
                                        Else        '//시작시간이후 출근
                                            STime = GetTime
                                        End If
                                        ETime = ToTime
                                    End If
                                                
                                End If
                            End If
                        Case "Y"    '//익일
                            If PosDate = OffDate Then '//기준일 = 퇴근일(당일퇴근)
                                STime = "0000"
                                ETime = "0000"
                            Else
                                If OffTime < FromTime Then '//퇴근시간 < 시작시간
                                    STime = "0000"
                                    ETime = "0000"
                                Else
                                    STime = FromTime
                                    If OffTime < ToTime Then '//퇴근시간 < 종료시간
                                        ETime = OffTime
                                    Else
                                        ETime = ToTime
                                    End If
                                End If
                            End If
                    End Select
'                Else
'                    If GNMUJO = "22" Then '//당일 익일 더해야 한다.
'                        Select Case NextDay
'                           Case "N" '//당일
'                            If PosDate <> GetDate Then '//당일 출근이 아니면
'                                STime = "0000"
'                                ETime = "0000"
'                            Else
'                                If PosDate = OffDate Then '//당일퇴근
'                                    If GetTime < FromTime Then  '//출근시간 < 시작시간
'                                        STime = FromTime  '//시작시간
'                                        If ToTime < OffTime Then    '//종료시간 < 퇴근시간
'                                            ETime = ToTime
'                                        Else
'                                            ETime = OffTime
'                                        End If
'                                    Else
'                                        If GetTime > ToTime Then    '//출근시간 > 종료시간
'                                            STime = "0000"
'                                            ETime = "0000"
'                                        Else
'                                            STime = GetTime    '//출근시간
'                                            If ToTime < OffTime Then    '//종료시간 < 퇴근시간
'                                                ETime = ToTime          '//종료시간
'                                            Else
'                                                ETime = OffTime         '//퇴근시간
'                                            End If
'                                        End If
'                                    End If
'                                Else '//다음날 퇴근
'                                    If ToTime < GetTime Then '//종료시간 < 출근시간
'                                        STime = "0000"
'                                        ETime = "0000"
'                                    Else
'                                        If GetTime < FromTime Then   '//출근시간 < 시작시간
'                                            STime = FromTime
'                                            If ToTime < OffTime Then '//종료시간 < 퇴근시간
'                                                ETime = ToTime
'                                            Else
'                                                ETime = OffTime
'                                            End If
'                                        Else        '//시작시간이후 출근
'                                            STime = GetTime
'                                            If ToTime < OffTime Then    '//종료시간 < 퇴근시간
'                                                ETime = ToTime          '//종료시간
'                                            Else
'                                                ETime = OffTime         '//퇴근시간
'                                            End If
'                                        End If
'                                    End If
'
'                                End If
'                            End If
'                        Case "Y"    '//익일
'                            If PosDate = OffDate Then '//기준일 = 퇴근일(당일퇴근)
'                                STime = "0000"
'                                ETime = "0000"
'                            Else
'                                If OffTime < FromTime Then '//퇴근시간 < 시작시간
'                                    STime = "0000"
'                                    ETime = "0000"
'                                Else
'                                    STime = FromTime
'                                    If OffTime < ToTime Then '//퇴근시간 < 종료시간
'                                        ETime = OffTime
'                                    Else
'                                        ETime = ToTime
'                                    End If
'                                End If
'                            End If
'                        End Select
                End If
                        
                hTime1 = PH_PY008_Time_Calc(STime, ETime) '//오전휴식시간

                If GNMUJO = "22" And TimeType = "65" Then '//야간조(2교대) 오전휴식일경우
                    Midnight = Midnight - hTime1
                Else
                    If DayOff = "1" Then
                        Base = Base - hTime1
                    Else
                        Special = Special - hTime1
                    End If
                End If
'            Case "15" '//점심휴식
'                If (GNMUJO = "11" Or GNMUJO = "21") Then '//1교대1조 2교대 1조
'                    If GetDate <> OffDate Then '다음날 퇴근
'                        If ToTime < GetTime Then '//출근시간이 점심시간이후이면
'                            STime = "0000"
'                            ETime = "0000"
'                        Else
'                            If FromTime < GetTime Then '//출근시간이 점심시작시간 이후이면
'                                STime = GetTime
'                                ETime = ToTime
'                            Else
'                                STime = FromTime
'                                ETime = ToTime
'                            End If
'                        End If
'                    Else '당일퇴근
'                        If ToTime < GetTime Then '//출근시간이 점심시간이후이면
'                            STime = "0000"
'                            ETime = "0000"
'                        Else
'                            If FromTime < GetTime Then
'                                STime = GetTime
'                                ETime = ToTime
'                            Else
'                                STime = FromTime
'                                If ToTime < OffTime Then
'                                    ETime = ToTime
'                                Else
'                                    ETime = OffTime
'                                End If
'                            End If
'                        End If
'                    End If
'                End If
'
'                hTime2 = PH_PY008_Time_Calc(STime, ETime) '//점심 시작시간
'
'
'                If DayOff = "1" Then
'                    Base = Base - hTime2
'                Else
'                    Special = Special - hTime2
'                End If
                
            Case "20", "60" '//연장근무
                If (GNMUJO = "11" Or GNMUJO = "21" Or GNMUJO = "31" Or GNMUJO = "32") Then '//1교대1조 2교대 1조
                    Select Case NextDay
                        Case "N" '//당일
                            If PosDate <> OffDate Then '// 출근일 <> 퇴근일
                                If GetTime < FromTime Then '//출근시간 < 시작시간
                                    STime = FromTime    '//
                                Else
                                    STime = GetTime
                                End If
                                ETime = "2400"
                            Else                            '//당일 퇴근
                                If OffTime < FromTime Then  '//퇴근시간 < 시작시간
                                    STime = "0000"
                                    ETime = "0000"
                                Else
                                    STime = FromTime    '//종료시간
                                    ETime = OffTime     '//퇴근시간
                                End If
                            End If
                            If DayOff = "1" Then
                                Extend = Extend + PH_PY008_Time_Calc(STime, ETime)
                            Else
                                SpExtend = SpExtend + PH_PY008_Time_Calc(STime, ETime)
                            End If
                        Case "Y" '//익일
                            If PosDate <> OffDate Then  '//기준일 <> 퇴근일
                                STime = "0000"
                                If OffTime < ToTime Then    '//퇴근시간 < 종료시간
                                    ETime = OffTime
                                Else
                                    ETime = ToTime
                                End If
                                
                            Else
                                STime = "0000"
                                ETime = "0000"
                            End If
                            If DayOff = "1" Then
                                Extend = Extend + PH_PY008_Time_Calc(STime, ETime)
                            Else
                                SpExtend = SpExtend + PH_PY008_Time_Calc(STime, ETime)
                            End If
                            
                    End Select
                ElseIf GNMUJO = "22" Or GNMUJO = "33" Then '//2교대 2조
                    Select Case NextDay
                        Case "N" '//당일
                            '//연장근무없다
                        Case "Y" '//익일
                            If PosDate <> OffDate Then '// 기준일 <> 퇴근일
                                If OffTime < FromTime Then  '//퇴근시간 < 시작시간
                                    STime = "0000"
                                    ETime = "0000"
                                Else
                                    STime = FromTime
                                    ETime = OffTime
                                    
                                End If
                            Else
                                STime = "0000"
                                ETime = "0000"
                            End If
                            
                            If DayOff = "1" Then
                                Extend = Extend + PH_PY008_Time_Calc(STime, ETime)
                            Else
                                SpExtend = SpExtend + PH_PY008_Time_Calc(STime, ETime)
                            End If
                    End Select
                End If
            Case "25" '//저녁휴식
            Case "30" '//심야시간
                If (GNMUJO = "11" Or GNMUJO = "21" Or GNMUJO = "22" Or GNMUJO = "31" Or GNMUJO = "32" Or GNMUJO = "33") Then '//1교대1조 2교대 1조
                    Select Case NextDay
                        Case "N" '//당일
                            If PosDate <> OffDate Then '//기준일자 <> 퇴근일자
                                If GetTime < FromTime Then '//출근시간 < 시작시간
                                    STime = FromTime    '//시작시간
                                Else
                                    STime = GetTime     '//출근시간
                                End If
                                ETime = "2400"
                            Else                            '//당일퇴근시
                                If OffTime < FromTime Then  '//퇴근시간 < 시작시간
                                    STime = "0000"
                                    ETime = "0000"
                                Else
                                    If GetTime < FromTime Then '//출근시간 < 시작시간
                                        STime = FromTime            '//시작시간
                                    Else
                                        STime = GetTime
                                    End If
                                    
                                    If OffTime < ToTime Then '//퇴근시간 < 종료시간
                                        ETime = OffTime
                                    Else
                                        ETime = ToTime '//종료시간
                                    End If
                                    
                                End If
                            End If
                            Midnight = Midnight + PH_PY008_Time_Calc(STime, ETime)
                        Case "Y" '//익일
                            If PosDate <> OffDate Then '//출근일 <> 퇴근일
                                STime = "0000"
                                If OffTime < ToTime Then    '//퇴근시간 < 종료시간
                                    ETime = OffTime         '//퇴근시간
                                Else
                                    ETime = ToTime          '//종료시간
                                End If
                            Else
                                STime = "0000"
                                ETime = "0000"
                            End If
                            Midnight = Midnight + PH_PY008_Time_Calc(STime, ETime)
                    End Select
'                ElseIf GNMUJO = "22" Then '//2교대 2조
'                    Select Case NextDay
'                        Case "N" '//당일
'                            If GetDate <> OffDate Then
'                                If GetTime < FromTime Then
'                                    STime = FromTime
'                                    ETime = "2400"
'                                Else
'                                    STime = GetTime
'                                    ETime = "2400"
'                                End If
'                            Else
'                                If GetTime < FromTime Then
'                                    STime = FromTime
'                                Else
'                                    STime = GetTime
'                                End If
'
'
'                                If OffTime < "2400" Then
'                                    ETime = OffTime
'                                Else
'                                    ETime = "2400"
'                                End If
'                            End If
'
'                            Midnight = Midnight + PH_PY008_Time_Calc(STime, ETime)
'                        Case "Y" '//익일
'                            If GetDate <> OffDate Then '// 출퇴근시간이 틀리면
'                                If OffTime < ToTime Then
'                                    STime = "0000"
'                                    ETime = OffTime
'                                Else
'                                    STime = "0000"
'                                    ETime = ToTime
'                                End If
'                            Else
'                                STime = "0000"
'                                ETime = "0000"
'                            End If
'
'                            Midnight = Midnight + PH_PY008_Time_Calc(STime, ETime)
'                    End Select
                End If
            Case "35" '//야간휴식 수정해야함
                '//hTime5
                Select Case NextDay
                    Case "N" '//당일
                        If PosDate = OffDate Then '//기준일자 = 퇴근일자
                            If OffTime < FromTime Then '//퇴근시간 < 시작시간
                                STime = "0000"
                                ETime = "0000"
                            Else
                                If GetTime < FromTime Then
                                    STime = FromTime
                                Else
                                    STime = GetTime
                                End If

                                If OffTime < ToTime Then
                                    ETime = OffTime
                                Else
                                    ETime = ToTime
                                End If
                            End If
                        Else
                            '//다음날 퇴근
                            If GetTime < FromTime Then '//출근시간 < 시작시간
                                STime = FromTime
                            Else
                                STime = GetTime
                            End If

                            ETime = ToTime
                        End If

                    Case "Y"
                        If PosDate = OffDate Then '//기준일자 = 퇴근일자
                            STime = "0000"
                            ETime = "0000"
                        Else
                            If OffTime < FromTime Then '//퇴근시간 < 시작시간
                                STime = "0000"
                                ETime = "0000"
                            Else
                                STime = FromTime
                                If OffTime < ToTime Then
                                    ETime = OffTime
                                Else
                                    ETime = ToTime
                                End If
                            End If
                        End If
                End Select

                hTime5 = PH_PY008_Time_Calc(STime, ETime)

                If DayOff = "1" Then '//평일
                    If GNMUJO = "22" Then '//2교대 2조(야간조)
                        Base = Base - hTime5 '//기본근무
                        Midnight = Midnight - hTime5 '//심야시간에서 차감
                    Else '//그외 주간조
                        Extend = Extend - hTime5 '//연장근무에서 차감
                        Midnight = Midnight - hTime5 '//심야시간에서 차감
                    End If
                Else
                    '//휴일
                    If GNMUJO = "22" Then
                        Special = Special - hTime5
                        Midnight = Midnight - hTime5 '//심야시간에서 차감
                    Else
                        '//다음날 퇴근은 연장근무임
                        SpExtend = SpExtend - hTime5 '//연장근무에서 차감
                        Midnight = Midnight - hTime5 '//심야시간에서 차감
                    End If
                End If
            
        End Select
        oRecordSet.MoveNext
    Next
    
    oForm.Items("EarlyTo").Specific.VALUE = PH_PY008_hhmm_Calc(EarlyTo) '//조출
    oForm.Items("Base").Specific.VALUE = PH_PY008_hhmm_Calc(Base, WorkType) '//기본
    oForm.Items("Extend").Specific.VALUE = PH_PY008_hhmm_Calc(Extend) '//연장
    oForm.Items("Midnight").Specific.VALUE = PH_PY008_hhmm_Calc(Midnight) '//심야
    
    oForm.Items("SEarlyTo").Specific.VALUE = PH_PY008_hhmm_Calc(SEarlyTo) '//특근조출
    oForm.Items("Special").Specific.VALUE = PH_PY008_hhmm_Calc(Special) '//특근
    oForm.Items("SpExtend").Specific.VALUE = PH_PY008_hhmm_Calc(SpExtend) '//특근연장
    
        
     
    
    Set oRecordSet = Nothing
    Exit Function
PH_PY008_Time_Calc_Main_Exit:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
PH_PY008_Time_Calc_Main_Error:
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "PH_PY008_Time_Calc_Main_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Private Function PH_PY008_Time_Calc(ByVal GetTime As String, ByVal OffTime As String) As Double
    
    Dim STime As Long
    Dim ETime As Long
    
    Dim hh As Long
    Dim MM As Double
    
    Dim Ret As Double
    
    
    STime = Int(Mid(GetTime, 1, 2)) * 3600 + Int(Mid(GetTime, 3, 2)) * 60
    
    ETime = Int(Mid(OffTime, 1, 2)) * 3600 + Int(Mid(OffTime, 3, 2)) * 60
    
'    hh = Int((ETime - STime) / 3600)
'    mm = ((ETime - STime) Mod 3600) / 60
'
'    If mm > 0 Then
'        If mm > 30 Then
'            mm = 1
'        Else
'            mm = 0.5
'        End If
'    ElseIf mm = 0 Then
'        mm = 0
'    Else
'        If mm < -30 Then
'            mm = -1
'        Else
'            mm = -0.5
'        End If
'    End If
'
'    Ret = hh + mm
    
    Ret = ETime - STime
    
    If IsNull(Ret) Then
        Ret = 0
    End If
    
    PH_PY008_Time_Calc = Ret
    

    
    Exit Function
PH_PY008_Time_Calct_Exit:
    Exit Function
PH_PY008_Mark_Set_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY008_Time_Calc_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY008_ChDate_Set()
    '//날자변경에 대한 자료 Set

    
    
    Dim i       As Long
    Dim sQry    As String
    Dim Day     As String
    
    Dim CLTCOD  As String
    Dim sPosDate As String
    
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY008_ChDate_Set_Error

    Call oForm.Freeze(True)
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    CLTCOD = oForm.Items("SCLTCOD").Specific.VALUE
    sPosDate = oForm.Items("SPosDate").Specific.VALUE
    
    '//해당일의 기본근태
    '//해당일의 근무일 등록 확인
    sQry = "Select U_WorkType From [@PH_PY003A] a Inner Join [@PH_PY003B] b On a.Code = b.Code "
    sQry = sQry + " Where a.U_CLTCOD = '" & CLTCOD & "' And b.U_Date = '" & sPosDate & "'"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
        Call MDC_Com.MDC_GF_Message("근태월력의 근무일 등록을 하지않습니다.", "E")
        GoTo PH_PY008_ChDate_Set_Exit
    Else
        oWorkType = oRecordSet.Fields(0).VALUE
        Set oCombo = oForm.Items("WorkType").Specific
        oCombo.Select oRecordSet.Fields(0).VALUE, psk_ByValue
        Call PH_PY008_ChGNMUJO_Set
    End If
    
    
    oForm.DataSources.UserDataSources.Item("PosDate").VALUE = sPosDate
    oForm.DataSources.UserDataSources.Item("GetDate").VALUE = sPosDate
    
    
'    oForm.Items("PosDate").Specific.Value = sPosDate
'    oForm.Items("GetDate").Specific.Value = sPosDate
    Day = PH_PY008_DaySelect(sPosDate)
                        
    oForm.DataSources.UserDataSources.Item("SDay").VALUE = Day
    oForm.DataSources.UserDataSources.Item("Day1").VALUE = Day
    
'    oForm.Items("SDay").Specific.Value = Day
'    oForm.Items("Day1").Specific.Value = Day
    Set oCombo = oForm.Items("SDayOff").Specific
    oCombo.Select PH_PY008_DayOffSelect(sPosDate), psk_ByValue
                         
    Set oCombo = oForm.Items("DayOff1").Specific
    oCombo.Select PH_PY008_DayOffSelect(sPosDate), psk_ByValue
                        
    
                        
    
    
'    Select Case oForm.Items("GNMUJO").Specific.Value
'        Case "22"
'            sQry = "Select DateAdd(dd, 1, '" & oForm.Items("SPosDate").Specific.Value & "')"
'            oRecordSet.DoQuery sQry
'            If oRecordSet.RecordCount > 0 Then
'                oForm.Items("OffDate").Specific.Value = Format(oRecordSet.Fields(0).Value, "YYYYMMDD")
'                oForm.Items("Day2").Specific.Value = PH_PY008_DaySelect(Format(oRecordSet.Fields(0).Value, "YYYYMMDD"))
'
'                Set oCombo = oForm.Items("DayOff2").Specific
'                oCombo.Select PH_PY008_DayOffSelect(Format(oRecordSet.Fields(0).Value, "YYYYMMDD")), psk_ByValue
'            End If
'
'        Case Else
'            oForm.Items("OffDate").Specific.Value = oForm.Items("SPosDate").Specific.Value
'            oForm.Items("Day2").Specific.Value = Day
'
'            Set oCombo = oForm.Items("DayOff2").Specific
'            oCombo.Select PH_PY008_DayOffSelect(oForm.Items.Item("SPosDate").Specific.Value), psk_ByValue
'
'    End Select
                            
    
                        
                        

    oForm.Update
    
    
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Sub
PH_PY008_ChDate_Set_Exit:
    Call oForm.Freeze(False)
    Exit Sub
PH_PY008_ChDate_Set_Error:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY008_ChDate_Set_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY008_ChGNMUJO_Set()
    '//근무조 변경에 대한 자료 Set

    
    Dim sQry    As String
    Dim Day     As String
    Dim Team1     As String
    Dim Team2    As String
    Dim Team3    As String
    Dim GNMUJO   As String
    
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY008_ChGNMUJO_Set_Error

    Call oForm.Freeze(True)
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    ' 근무조를 받아와서 그 근무조가 야간인지 주간인지 휴무인지 체크함. (주52)
    ' 근무조를 받아와서 체크해서 21 or 22로 변경하는 건 시간표(PH_PY003) 때문임.
    Team1 = Trim(oForm.Items("Team1").Specific.VALUE)
    Team2 = Trim(oForm.Items("Team2").Specific.VALUE)
    Team3 = Trim(oForm.Items("Team3").Specific.VALUE)
    GNMUJO = Trim(oForm.Items("GNMUJO").Specific.VALUE)
    If (GNMUJO = "41") Then
        If (Team1 = "D") Then
             GNMUJO = "21"
        ElseIf (Team1 = "N") Then
             GNMUJO = "22"
        Else
             GNMUJO = ""
        End If
    End If
        
    If (GNMUJO = "42") Then
        If (Team2 = "D") Then
             GNMUJO = "21"
        ElseIf (Team2 = "N") Then
             GNMUJO = "22"
        Else
             GNMUJO = ""
        End If
    End If
            
    If (GNMUJO = "43") Then
        If (Team3 = "D") Then
             GNMUJO = "21"
        ElseIf (Team3 = "N") Then
             GNMUJO = "22"
        Else
             GNMUJO = ""
        End If
    End If
                        
    Select Case GNMUJO
        Case "22", "33"
            sQry = "Select DateAdd(dd, 1, '" & oForm.Items("SPosDate").Specific.VALUE & "')"
            oRecordSet.DoQuery sQry
            If oRecordSet.RecordCount > 0 Then
                oForm.DataSources.UserDataSources.Item("OffDate").VALUE = Format(oRecordSet.Fields(0).VALUE, "YYYYMMDD")
                oForm.DataSources.UserDataSources.Item("Day2").VALUE = PH_PY008_DaySelect(Format(oRecordSet.Fields(0).VALUE, "YYYYMMDD"))
                
'                oForm.Items("OffDate").Specific.Value = Format(oRecordset.Fields(0).Value, "YYYYMMDD")
'                oForm.Items("Day2").Specific.Value = PH_PY008_DaySelect(Format(oRecordset.Fields(0).Value, "YYYYMMDD"))
                                    
                Set oCombo = oForm.Items("DayOff2").Specific
                oCombo.Select PH_PY008_DayOffSelect(Format(oRecordSet.Fields(0).VALUE, "YYYYMMDD")), psk_ByValue
            End If
                                
        Case Else
            oForm.DataSources.UserDataSources.Item("OffDate").VALUE = oForm.Items("SPosDate").Specific.VALUE
            oForm.DataSources.UserDataSources.Item("Day2").VALUE = PH_PY008_DaySelect(oForm.Items.Item("SPosDate").Specific.VALUE)
            
'            oForm.Items("OffDate").Specific.Value = oForm.Items("SPosDate").Specific.Value
'            oForm.Items("Day2").Specific.Value = PH_PY008_DaySelect(oForm.Items.Item("SPosDate").Specific.Value)
                                 
            Set oCombo = oForm.Items("DayOff2").Specific
            oCombo.Select PH_PY008_DayOffSelect(oForm.Items.Item("SPosDate").Specific.VALUE), psk_ByValue
                            
    End Select
    
    If oForm.Items("GNMUJO").Specific.VALUE = "11" And oForm.Items("SDayOff").Specific.VALUE = "1" Then
        '//근무조가 1교대1조 평일일 경우
        Select Case oForm.Items("SCLTCOD").Specific.VALUE
            Case "1", "3"
                oForm.Items("GetTime").Specific.VALUE = "0830"
                oForm.Items("OffTime").Specific.VALUE = "1730"
'                oForm.DataSources.UserDataSources.Item("GetTime").Value = "0830"
'                oForm.DataSources.UserDataSources.Item("OffTime").Value = "1730"
                
            Case "2"
                oForm.Items("GetTime").Specific.VALUE = "0750"
                oForm.Items("OffTime").Specific.VALUE = "1700"
            
'                oForm.DataSources.UserDataSources.Item("GetTime").Value = "0750"
'                oForm.DataSources.UserDataSources.Item("OffTime").Value = "1700"
        End Select
    Else
        oForm.Items("GetTime").Specific.VALUE = "0000"
        oForm.Items("OffTime").Specific.VALUE = "0000"
    
'        oForm.DataSources.UserDataSources.Item("GetTime").Value = "0000"
'        oForm.DataSources.UserDataSources.Item("OffTime").Value = "0000"
    End If
'    oForm.Items("GetTime").Specific.Value = "0000"
'    oForm.Items("OffTime").Specific.Value = "0000"
    
    
    
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Sub
PH_PY008_ChGNMUJO_Set_Exit:
    Call oForm.Freeze(False)
    Exit Sub
PH_PY008_ChGNMUJO_Set_Error:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY008_ChGNMUJO_Set_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY008_Time_ReSet()
    '//시간 초기화
    Call oForm.Freeze(True)
    oForm.DataSources.UserDataSources.Item("Base").VALUE = 0
    oForm.DataSources.UserDataSources.Item("Extend").VALUE = 0
    oForm.DataSources.UserDataSources.Item("Midnight").VALUE = 0
    oForm.DataSources.UserDataSources.Item("EarlyTo").VALUE = 0
    oForm.DataSources.UserDataSources.Item("Special").VALUE = 0
    oForm.DataSources.UserDataSources.Item("SpExtend").VALUE = 0
    oForm.DataSources.UserDataSources.Item("EducTran").VALUE = 0
    oForm.DataSources.UserDataSources.Item("SEarlyTo").VALUE = 0
    oForm.DataSources.UserDataSources.Item("LateTo").VALUE = 0
    oForm.DataSources.UserDataSources.Item("EarlyOff").VALUE = 0
    oForm.DataSources.UserDataSources.Item("GoOut").VALUE = 0

    oForm.DataSources.UserDataSources.Item("GOFrTime").VALUE = "0000"
    oForm.DataSources.UserDataSources.Item("GOToTime").VALUE = "0000"
    oForm.DataSources.UserDataSources.Item("GOFrTim2").VALUE = "0000"
    oForm.DataSources.UserDataSources.Item("GOToTim2").VALUE = "0000"
    

'                    oForm.Items("MSTCOD").Specific.Value = ""
'                    oForm.Items("FullName").Specific.Value = ""
'                    oForm.Items("GetTime").Specific.Value = "0000"
'                    oForm.Items("OffTime").Specific.Value = "0000"
'
''                    Set oCombo = oForm.Items("ShiftDat").Specific
''                    oCombo.Select "", psk_ByValue
''
''                    Set oCombo = oForm.Items("GNMUJO").Specific
''                    oCombo.Select "", psk_ByValue
'
'                    oForm.Items("ActCode").Specific.Value = ""
'                    oForm.Items("ActText").Specific.Value = ""
'
'                    Set oCombo = oForm.Items("DangerCD").Specific
'                    oCombo.Select "", psk_ByValue
'
'                    oForm.Items("RToDate").Specific.Value = ""
'                    oForm.Items("ROffDate").Specific.Value = ""
'                    oForm.Items("RToTime").Specific.Value = "0000"
'                    oForm.Items("ROffTime").Specific.Value = "0000"
'
'                    Set oCombo = oForm.Items("Confirm").Specific
'                    oCombo.Select "N", psk_ByValue

    
    Call oForm.Freeze(False)
    Exit Sub
PH_PY008_Time_ReSet_Exit:
    Call oForm.Freeze(False)
    Exit Sub
PH_PY008_Time_ReSet_Error:
    
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY008_Time_ReSet_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PH_PY008_OverTimeCheck(ByVal PosDate As String, ByVal MSTCOD As String) As Double

    '//주 근무시간
    
    Dim sQry    As String
    
    Dim FrDate As String
    Dim ToDate As String
    Dim CLTCOD As String
    Dim sunday As String
    Dim OverTimeYN As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY008_OverTimeCheck_Error

    Call oForm.Freeze(True)
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    CLTCOD = oForm.Items("SCLTCOD").Specific.VALUE
    
    If oForm.Items("SDay").Specific.VALUE = "일" Then
        
    
        
        ToDate = PosDate
        
        '//월요일 찾기
        sQry = " SELECT U_Date "
        sQry = sQry + " From [@PH_PY003A] a Inner join [@PH_PY003B] b On a.Code = b.Code And a.U_CLTCOD = '" & CLTCOD & "'"
        sQry = sQry + " Where DatePart(Weekday, U_Date) = 2 "
        sQry = sQry + " and U_Date between DATEADD(dd, -6, '" & PosDate & "') and '" & PosDate & "'"
        
        oRecordSet.DoQuery sQry
        
        If oRecordSet.RecordCount > 0 Then
            FrDate = oRecordSet.Fields(0).VALUE
        Else
            Call MDC_Com.MDC_GF_Message("[12시간초과체크] 기준일자(월요일)을 가져오지 못했습니다..", "E")
            GoTo PH_PY008_OverTimeCheck_Exit
        End If
        
    ElseIf oForm.Items("SDay").Specific.VALUE = "월" Then
        FrDate = PosDate
        
        '//일요일 찾기
        sQry = " SELECT U_Date "
        sQry = sQry + " From [@PH_PY003A] a Inner join [@PH_PY003B] b On a.Code = b.Code And a.U_CLTCOD = '" & CLTCOD & "'"
        sQry = sQry + " Where DatePart(Weekday, U_Date) = 1 "
        sQry = sQry + " and U_Date between '" & PosDate & "' and  DATEADD(dd, 6, '" & PosDate & "')"
        
        oRecordSet.DoQuery sQry
        If oRecordSet.RecordCount > 0 Then
            ToDate = oRecordSet.Fields(0).VALUE
        Else
            Call MDC_Com.MDC_GF_Message("[12시간초과체크] 기준일자(일요일)을 가져오지 못했습니다..", "E")
            GoTo PH_PY008_OverTimeCheck_Exit
        End If
    Else
        '//월요일 찾기
        sQry = " SELECT U_Date "
        sQry = sQry + " From [@PH_PY003A] a Inner join [@PH_PY003B] b On a.Code = b.Code And a.U_CLTCOD = '" & CLTCOD & "'"
        sQry = sQry + " Where DatePart(Weekday, U_Date) = 2 "
        sQry = sQry + " and U_Date between DATEADD(dd, -6, '" & PosDate & "') and '" & PosDate & "'"
        
        oRecordSet.DoQuery sQry
        If oRecordSet.RecordCount > 0 Then
            FrDate = oRecordSet.Fields(0).VALUE
        Else
            Call MDC_Com.MDC_GF_Message("[12시간초과체크] 기준일자(월요일)을 가져오지 못했습니다..", "E")
            GoTo PH_PY008_OverTimeCheck_Exit
        End If
        
        '//일요일 찾기
        sQry = " SELECT U_Date "
        sQry = sQry + " From [@PH_PY003A] a Inner join [@PH_PY003B] b On a.Code = b.Code And a.U_CLTCOD = '" & CLTCOD & "'"
        sQry = sQry + " Where DatePart(Weekday, U_Date) = 1 "
        sQry = sQry + " and U_Date between '" & PosDate & "' and  DATEADD(dd, 6, '" & PosDate & "')"
        
        oRecordSet.DoQuery sQry
        If oRecordSet.RecordCount > 0 Then
            ToDate = oRecordSet.Fields(0).VALUE
        Else
            Call MDC_Com.MDC_GF_Message("[12시간초과체크] 기준일자(일요일)을 가져오지 못했습니다..", "E")
            GoTo PH_PY008_OverTimeCheck_Exit
        End If
        
    End If
        
    sQry = "SELECT Isnull(U_Inform05,'N') From [@PH_PY001A] Where Code = '" & MSTCOD & "'"
    oRecordSet.DoQuery sQry

    OverTimeYN = oRecordSet.Fields(0).VALUE
    
    ' 기존 로직에서 Base Special 추가로 입력해서 주 52시간 초과인지 계산함.
    sQry = " Select Sum(Base + Special + Extend + SpExtend + EarlyTo + SEarlyTo) From ZPH_PY008 Where PosDate Between '" & FrDate & "' and '" & ToDate & "'"
    sQry = sQry + " And MSTCOD = '" & MSTCOD & "'"
    sQry = sQry + " And PosDate <> '" & PosDate & "'"
    oRecordSet.DoQuery sQry
    
    If (oRecordSet.RecordCount = 0) Then
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
        GoTo PH_PY008_OverTimeCheck_Exit
    Else
        If OverTimeYN = "Y" Then
            PH_PY008_OverTimeCheck = -100
        Else
            PH_PY008_OverTimeCheck = oRecordSet.Fields(0).VALUE
        End If
    End If
    
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
PH_PY008_OverTimeCheck_Exit:
    PH_PY008_OverTimeCheck = 15 '//주계산 시간
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
PH_PY008_OverTimeCheck_Error:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY008_OverTimeCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY008_Delete()
    '//선택된 자료 삭제
    Dim PosDate As String
    Dim CLTCOD As String
    Dim MSTCOD As String
    Dim FullName As String
    
    Dim i As Integer
    Dim sQry    As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY008_Delete_Error

    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    
    
    Call oForm.Freeze(True)
    
    CLTCOD = oForm.Items("SCLTCOD").Specific.VALUE
    
    MSTCOD = oForm.Items("MSTCOD").Specific.VALUE
    FullName = oForm.Items("FullName").Specific.VALUE
    PosDate = oForm.Items("PosDate").Specific.VALUE
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    
    
    
'    If Trim(MSTCOD) <> "" Then
'        '//선택한 한명만 삭제
'        If Sbo_Application.MessageBox(" 선택한사원('" & FullName & "')을 삭제하시겠습니까? ?", "2", "예", "아니오") = "1" Then
'            sQry = "Delete From ZPH_PY008 Where CLTCOD = '" & CLTCOD & "' AND  PosDate = '" & PosDate & "' And MSTCOD = '" & MSTCOD & "' "
'            oRecordSet.DoQuery sQry
'        End If
'    Else
        If Sbo_Application.MessageBox(" 선택한사원 전체를 삭제하시겠습니까? ?", "2", "예", "아니오") = "1" Then
            If oDS_PH_PY008.Rows.Count > 0 Then
                Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("삭제시작!", oDS_PH_PY008.Rows.Count, False)
                
                For i = 0 To oDS_PH_PY008.Rows.Count - 1
                    If oDS_PH_PY008.Columns.Item("Chk").Cells(i).VALUE = "Y" Then
                        
                        PosDate = Trim(oDS_PH_PY008.Columns.Item("GetDate").Cells(i).VALUE)
                        MSTCOD = oDS_PH_PY008.Columns.Item("MSTCOD").Cells(i).VALUE
                        
                        sQry = "Delete From ZPH_PY008 Where CLTCOD = '" & CLTCOD & "' AND  PosDate = '" & PosDate & "' And MSTCOD = '" & MSTCOD & "' "
                        oRecordSet.DoQuery sQry
                    End If
                    ProgressBar01.VALUE = ProgressBar01.VALUE + 1
                    ProgressBar01.Text = ProgressBar01.VALUE & "/" & oDS_PH_PY008.Rows.Count & "건 삭제중...!"
                Next
            End If
        Else
            GoTo PH_PY008_Delete_Exit
        End If
'    End If
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
        Set ProgressBar01 = Nothing
    End If
    
    Call PH_PY008_MTX01("A")
    
    Call oForm.Freeze(False)
    
    Set oRecordSet = Nothing
    
    Exit Sub
PH_PY008_Delete_Exit:
    Set oRecordSet = Nothing
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
        Set ProgressBar01 = Nothing
    End If
    
    Call oForm.Freeze(False)
    Exit Sub
PH_PY008_Delete_Error:
    Set oRecordSet = Nothing
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
        Set ProgressBar01 = Nothing
    End If
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY008_Delete_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PH_PY008_Extend15(ByVal MSTCOD As String) As Double

    '//사무기술직 월 15시간 => 20시간으로 변경 2014/07/01부 N.G.Y
    
    Dim sQry    As String
    
    
    Dim CLTCOD As String
    Dim YM As String
    
    Dim StTime As Double
    
    Dim OverTime As Double
    
    Dim CLTCODTIME As Double
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY008_Extend15_Error

    Call oForm.Freeze(True)
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    
    
    OverTime = PH_PY008_OverTimeCheck(oForm.Items("PosDate").Specific.VALUE, MSTCOD)
    
    CLTCOD = oForm.Items("SCLTCOD").Specific.VALUE
    YM = Left(oForm.Items("PosDate").Specific.VALUE, 6)
        
        
    If CLTCOD = "1" Then
        CLTCODTIME = 2.5
    Else
        CLTCODTIME = 2
    End If
    
    If OverTime + CLTCODTIME > 12 Then  '주 12시간 초과되면 안됨
        CLTCODTIME = 0
    End If
    
    If oForm.Items("SDayOff").Specific.VALUE = "1" Then '//평일
        StTime = 20 '//월 20시간
    
        sQry = " Select Isnull(Sum(Extend),0)"
        sQry = sQry + " From ZPH_PY008"
        sQry = sQry + " Where Convert(char(6),PosDate,112) = '" & YM & "'"
        sQry = sQry + " And CLTCOD = '" & CLTCOD & "'"
        sQry = sQry + " AND MSTCOD = '" & MSTCOD & "'"
        
        oRecordSet.DoQuery sQry
        
        If oRecordSet.RecordCount > 0 Then
            If oRecordSet.Fields(0).VALUE - StTime < 0 And Abs(oRecordSet.Fields(0).VALUE - StTime) >= CLTCODTIME Then
                PH_PY008_Extend15 = CLTCODTIME
            ElseIf oRecordSet.Fields(0).VALUE - StTime >= 0 Then
                PH_PY008_Extend15 = 0
            Else
                PH_PY008_Extend15 = Abs(oRecordSet.Fields(0).VALUE - StTime)
            End If
        Else
            Call MDC_Com.MDC_GF_Message("사무기술직(월급제)의 연장근무시간 계산이 되지않았습니다. .", "E")
            GoTo PH_PY008_Extend15_Exit
        End If
    End If
    
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
PH_PY008_Extend15_Exit:
    PH_PY008_Extend15 = 0
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
PH_PY008_Extend15_Error:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY008_Extend15_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Function PH_PY008_hhmm_Calc(ByVal mTime As Double, Optional ByVal pWorkType As String) As Double
    
    Dim hh As Long
    Dim MM As Double
    
    Dim Ret As Double
    
    hh = Int((mTime) / 3600)
    MM = ((mTime) Mod 3600) / 60

    If MM > 0 Then
        If MM > 30 Then
            MM = 1
        Else
            MM = 0.5
        End If
    ElseIf MM = 0 Then
        MM = 0
    Else
        If MM < -30 Then
            MM = -1
        Else
            MM = -0.5
        End If
    End If

    Ret = hh + MM
    '근태구분이 반차일 경우 무조건 4시간 반환(2014.04.21 송명규 추가) 별로 안 좋은 방법인데...
    If pWorkType = "D09" Or pWorkType = "D10" Then
        Ret = 4
    End If
        
    If IsNull(Ret) Then
        Ret = 0
    End If
    
    PH_PY008_hhmm_Calc = Ret
    
    Exit Function
PH_PY008_hhmm_Calc_Exit:
    Exit Function
PH_PY008_hhmm_Calc_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY008_hhmm_Calc_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY008_TitleSetting(iRow As Long)
    Dim i               As Long
    Dim j               As Long
    Dim sQry            As String
     
    Dim COLNAM(29)       As String

    Dim oColumn         As SAPbouiCOM.EditTextColumn
    Dim oComboCol       As SAPbouiCOM.ComboBoxColumn
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
     
    On Error GoTo Error_Message
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    oForm.Freeze True
    
    COLNAM(0) = "순번"
    COLNAM(1) = "선택"
    COLNAM(2) = "부서"
    COLNAM(3) = "담당"
    COLNAM(4) = "반"
    COLNAM(5) = "사번"
    COLNAM(6) = "성명"
    COLNAM(7) = "요일"
    COLNAM(8) = "휴일"
    COLNAM(9) = "기본"
    COLNAM(10) = "연장"
    COLNAM(11) = "심야"
    COLNAM(12) = "특근"
    COLNAM(13) = "특연"
    COLNAM(14) = "조출"
    COLNAM(15) = "휴조"
    COLNAM(16) = "교육"
    COLNAM(17) = "지각"
    COLNAM(18) = "조퇴"
    COLNAM(19) = "외출"
    COLNAM(20) = "위해"
    COLNAM(21) = "직무"
    COLNAM(22) = "근태"
    COLNAM(23) = "교대일수"
    COLNAM(24) = "출근일자"
    COLNAM(25) = "출근시간"
    COLNAM(26) = "퇴근일자"
    COLNAM(27) = "퇴근시간"
    COLNAM(28) = "비고"
    COLNAM(29) = "교대인정대상"

    For i = 0 To UBound(COLNAM)
        oGrid1.Columns.Item(i).TitleObject.Caption = COLNAM(i)
        oGrid1.Columns.Item(i).Editable = False
        

        Select Case COLNAM(i)
            Case "선택"
                oGrid1.Columns.Item(i).Editable = True
                oGrid1.Columns.Item(i).Type = gct_CheckBox
            Case "부서"
                oGrid1.Columns.Item(i).Type = gct_ComboBox
                Set oComboCol = oGrid1.Columns.Item("TeamCode")

                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
                sQry = sQry & " WHERE Code = '1' AND U_UseYN= 'Y' Order by U_Seq"
                oRecordSet.DoQuery sQry
                oComboCol.ValidValues.Add "", ""
                If oRecordSet.RecordCount > 0 Then
                    For j = 0 To oRecordSet.RecordCount - 1
                        oComboCol.ValidValues.Add oRecordSet.Fields(0).VALUE, oRecordSet.Fields(1).VALUE
                        oRecordSet.MoveNext
                    Next j
                End If


                oComboCol.DisplayType = cdt_Description
            Case "담당"
                oGrid1.Columns.Item(i).Type = gct_ComboBox
                Set oComboCol = oGrid1.Columns.Item("RspCode")

                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
                sQry = sQry & " WHERE Code = '2' AND U_UseYN= 'Y' Order by U_Seq"
                oRecordSet.DoQuery sQry
                oComboCol.ValidValues.Add "", ""
                If oRecordSet.RecordCount > 0 Then
                    For j = 0 To oRecordSet.RecordCount - 1
                        oComboCol.ValidValues.Add oRecordSet.Fields(0).VALUE, oRecordSet.Fields(1).VALUE
                        oRecordSet.MoveNext
                    Next j
                End If


                oComboCol.DisplayType = cdt_Description
            Case "반"
                oGrid1.Columns.Item(i).Type = gct_ComboBox
                Set oComboCol = oGrid1.Columns.Item("ClsCode")

                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
                sQry = sQry & " WHERE Code = '9' AND U_UseYN= 'Y' Order by U_Seq"
                oRecordSet.DoQuery sQry
                oComboCol.ValidValues.Add "", ""
                If oRecordSet.RecordCount > 0 Then
                    For j = 0 To oRecordSet.RecordCount - 1
                        oComboCol.ValidValues.Add oRecordSet.Fields(0).VALUE, oRecordSet.Fields(1).VALUE
                        oRecordSet.MoveNext
                    Next j
                End If


                oComboCol.DisplayType = cdt_Description
            Case "휴일"
                oGrid1.Columns.Item(i).Type = gct_ComboBox
                Set oComboCol = oGrid1.Columns.Item("DayOff")

                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P202'"
                oRecordSet.DoQuery sQry
                oComboCol.ValidValues.Add "", ""
                If oRecordSet.RecordCount > 0 Then
                    For j = 0 To oRecordSet.RecordCount - 1
                        oComboCol.ValidValues.Add oRecordSet.Fields(0).VALUE, oRecordSet.Fields(1).VALUE
                        oRecordSet.MoveNext
                    Next j
                End If


                oComboCol.DisplayType = cdt_Description
            Case "근태"
                oGrid1.Columns.Item(i).Type = gct_ComboBox
                Set oComboCol = oGrid1.Columns.Item("WorkType")
                
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
                sQry = sQry & " WHERE Code = 'P221' AND U_UseYN= 'Y' Order by U_Seq"
                oRecordSet.DoQuery sQry
                If oRecordSet.RecordCount > 0 Then
                    For j = 0 To oRecordSet.RecordCount - 1
                        oComboCol.ValidValues.Add oRecordSet.Fields(0).VALUE, oRecordSet.Fields(1).VALUE
                        oRecordSet.MoveNext
                    Next j
               End If
    
                oComboCol.DisplayType = cdt_Description
        
            Case "기본", "연장", "심야", "특근", "특연", "조출", "휴조", "교육", "지각", "조퇴", "외출"
                oGrid1.Columns.Item(i).RightJustified = True

            Case Else
                oGrid1.Columns.Item(i).Editable = False
        End Select
        
    Next i
    
    oGrid1.AutoResizeColumns

    oForm.Freeze False

    Set oColumn = Nothing

    Exit Sub

Error_Message:
    oForm.Freeze False
    Set oColumn = Nothing
    Sbo_Application.SetStatusBarMessage "PH_PY008_TitleSetting Error : " & Space(10) & Err.Description, bmt_Short, True
End Sub


Private Function PH_PY008_ModifyYN() As String
On Error GoTo PH_PY008_ModifyYN_Error
    '//추가 수정 삭제 가능한지
    
    Dim Today       As String
    Dim Todaytm     As String
    
    Dim sQry    As String

    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '//안강사업장은 제외
    If oForm.Items("SCLTCOD").Specific.VALUE <> "3" Then
    
        Today = Format(Now, "yyyymmdd")
        Todaytm = Format(Now, "HHMM")
                        
        If oForm.Items("PosDate").Specific.VALUE < Today Then
            If oRspCodeYN = "Y" Then '담당권한이 있으면...
               PH_PY008_ModifyYN = "Y"
            Else
               PH_PY008_ModifyYN = "N"
               Sbo_Application.SetStatusBarMessage "이전일자는 추가/수정/삭제를 할수 없습니다. 관리담당으로 확인바랍니다.", bmt_Short, True
            End If
        ElseIf oForm.Items("PosDate").Specific.VALUE = Today Then
            If Todaytm > "1500" Then
                If oRspCodeYN = "Y" Then '담당권한이 있으면...
                    PH_PY008_ModifyYN = "Y"
                Else
                    PH_PY008_ModifyYN = "N"
                    Sbo_Application.SetStatusBarMessage "15시 이후 추가/수정/삭제를 할수 없습니다. 관리담당으로 확인바랍니다.", bmt_Short, True
                End If
            Else
                PH_PY008_ModifyYN = "Y"
            End If
        Else
            PH_PY008_ModifyYN = "Y"
        End If
    Else
        PH_PY008_ModifyYN = "Y"
    End If
    
    
    Set oRecordSet = Nothing
    Exit Function
PH_PY008_ModifyYN_Exit:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
PH_PY008_ModifyYN_Error:
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "PH_PY008_ModifyYN_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Function PH_PY008_Special(ByVal MSTCOD As String) As Double

    '//사무기술직 특근시간 월 24시간 초과 할 수 없음. 2014/07/01부 N.G.Y
    
    Dim sQry    As String
    
    
    Dim CLTCOD As String
    Dim YM As String
    
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY008_Special_Error

    Call oForm.Freeze(True)
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
        
    CLTCOD = oForm.Items("SCLTCOD").Specific.VALUE
    YM = Left(oForm.Items("PosDate").Specific.VALUE, 6)
    
    sQry = " Select Isnull(Sum(Special),0)"
    sQry = sQry + " From ZPH_PY008"
    sQry = sQry + " Where Convert(char(6),PosDate,112) = '" & YM & "'"
    sQry = sQry + " And CLTCOD = '" & CLTCOD & "'"
    sQry = sQry + " AND MSTCOD = '" & MSTCOD & "'"
    
    oRecordSet.DoQuery sQry
        
    If oRecordSet.RecordCount > 0 Then
       PH_PY008_Special = oRecordSet.Fields(0).VALUE
    Else
       PH_PY008_Special = 0
    End If
    
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
PH_PY008_Special_Exit:
    PH_PY008_Special = 0
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
PH_PY008_Special_Error:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY008_Special_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

'Private Sub PH_PY008_View()
'
'    '//메트릭스에 근태이상자 로드
'
'    Dim i       As Long
'    Dim sQry    As String
'    Dim iRow        As Long
'
'    Dim PosDate As String
'    Dim CLTCOD As String
'
'    Dim oRecordSet As SAPbobsCOM.Recordset
'
'    On Error GoTo PH_PY008_View_Error
'
'    Call oForm.Freeze(True)
'
'    oMat1.Clear
''    oMat1.FlushToDataSource
''    oMat1.LoadFromDataSource
'
'    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
'
'    PosDate = oForm.Items("PosDate").Specific.Value
'    CLTCOD = oForm.Items("SCLTCOD").Specific.Value
'
'    sQry = "select PosDate = Convert(Char(8),p.PosDate,112),"
'    sQry = sQry + " TeamName = (SELECT U_CodeNm FROM [@PS_HR200L] where Code = '1' and  U_Code = p.TeamCode ),"
'    sQry = sQry + " RspName = (SELECT U_CodeNm FROM [@PS_HR200L] where Code = '2' and  U_Code = p.RspCode ),"
'    sQry = sQry + " MSTCOD = p.MSTCOD, p.FullName"
'    sQry = sQry + " from ZPH_PY008 p"
'    sQry = sQry + " Where p.PosDate    = '" & PosDate & "'"
'    sQry = sQry + " And p.CLTCOD = '" & CLTCOD & "'"
'    sQry = sQry + " and p.Confirm  = 'Y'"
'    sQry = sQry + " and p.Attend = 'N'"
'
'    oRecordSet.DoQuery sQry
'    If oRecordSet.RecordCount = 0 Then
'        Call MDC_Com.MDC_GF_Message("근태이상자 자료가 없습니다..", "E")
'        GoTo PH_PY008_View_Exit
'    End If
'
'    For i = 1 To oRecordSet.RecordCount
'        oMat1.AddRow (1)
'        oMat1.Columns("LineNum").Cells(i).Specific.Value = i
'        oMat1.Columns("PosDate").Cells(i).Specific.Value = oRecordSet.Fields(0).Value
'        oMat1.Columns("TeamName").Cells(i).Specific.Value = oRecordSet.Fields(1).Value
'        oMat1.Columns("RspName").Cells(i).Specific.Value = oRecordSet.Fields(2).Value
'        oMat1.Columns("MSTCOD").Cells(i).Specific.Value = oRecordSet.Fields(3).Value
'        oMat1.Columns("FullName").Cells(i).Specific.Value = oRecordSet.Fields(4).Value
'
''        If i <> 0 Then
''            oDS_PH_PYMat.InsertRecord (i)
''        End If
''        oDS_PH_PYMat.Offset = i
''        Call oDS_PH_PYMat.setValue("U_LineNum", i, i + 1)
''        Call oDS_PH_PYMat.setValue("U_ColReg01", i, oRecordSet.Fields(0).Value)
''        Call oDS_PH_PYMat.setValue("U_ColReg02", i, oRecordSet.Fields(1).Value)
''        Call oDS_PH_PYMat.setValue("U_ColReg03", i, oRecordSet.Fields(2).Value)
''        Call oDS_PH_PYMat.setValue("U_ColReg04", i, oRecordSet.Fields(3).Value)
''        Call oDS_PH_PYMat.setValue("U_ColReg05", i, oRecordSet.Fields(4).Value)
'        oRecordSet.MoveNext
''
'
'    Next
'
''    oMat1.LoadFromDataSource
'    oMat1.AutoResizeColumns
'    oForm.Items("Mat01").Visible = True
'    oForm.Update
'
'    Set oRecordSet = Nothing
'    Call oForm.Freeze(False)
'    Exit Sub
'PH_PY008_View_Exit:
'    Set oRecordSet = Nothing
'    Call oForm.Freeze(False)
'    Exit Sub
'PH_PY008_View_Error:
'    Set oRecordSet = Nothing
'    Call oForm.Freeze(False)
'    Sbo_Application.SetStatusBarMessage "PH_PY008_View_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
'End Sub
