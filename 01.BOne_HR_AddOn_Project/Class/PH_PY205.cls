VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY205"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File               : PH_PY205.cls
'//  Module             : 인사관리>인사관리>인사정보관리
'//  Desc               : 교육계획VS실적조회
'//  FormType           : PH_PY205
'//  Create Date(Start) : 2016.12.23
'//  Create Date(End)   :
'//  Creator            : Song Myoung gyu
'//  Modified Date      :
'//  Modifier           :
'//  Company            : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form

Public oGrid01 As SAPbouiCOM.Grid
Public oGrid02 As SAPbouiCOM.Grid
Public oGrid03 As SAPbouiCOM.Grid

Public oDS_PH_PY205A As SAPbouiCOM.DataTable
Public oDS_PH_PY205B As SAPbouiCOM.DataTable
Public oDS_PH_PY205C As SAPbouiCOM.DataTable

Private oLastItemUID As String
Private oLastColUID  As String
Private oLastColRow  As Long

Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error

    Dim i As Long
    Dim oXmlDoc As New MSXML2.DOMDocument
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY205.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY205_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY205"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
        
    Call oForm.Freeze(True)
    Call PH_PY205_CreateItems
    Call PH_PY205_ComboBox_Setting
    Call PH_PY205_EnableMenus
    Call oForm.Items("Folder01").Specific.Select '폼이 로드 될 때 Folder01이 선택됨
    Call PH_PY205_SetDocument(oFromDocEntry01)
    'Call PH_PY205_FormResize
    
    Call oForm.Update
    Call oForm.Freeze(False)

    oForm.DataSources.UserDataSources("StdYear01").VALUE = Format(Now, "YYYY") '기준년도
    oForm.DataSources.UserDataSources.Item("CancelYN01").VALUE = "Y"
    
    oForm.DataSources.UserDataSources.Item("FrMt02").VALUE = Format(Now, "YYYYMM") '기준년월(시작)
    oForm.DataSources.UserDataSources.Item("ToMt02").VALUE = Format(Now, "YYYYMM") '기준년월(종료)
    oForm.DataSources.UserDataSources.Item("Check201").VALUE = "Y" '수정계획 미출력

    oForm.DataSources.UserDataSources("StdYear03").VALUE = Format(Now, "YYYY") '기준년도
    oForm.DataSources.UserDataSources.Item("CancelYN03").VALUE = "Y"

    Call oForm.Items("MSTCOD01").CLICK(ct_Regular)

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

Private Function PH_PY205_CreateItems() As Boolean
On Error GoTo PH_PY205_CreateItems_Error

    Dim sQry As String
    Dim i As Long
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Set oGrid01 = oForm.Items("Grid01").Specific
    Set oGrid02 = oForm.Items("Grid02").Specific
    Set oGrid03 = oForm.Items("Grid03").Specific
    
    Call oForm.DataSources.DataTables.Add("PH_PY205A")
    Call oForm.DataSources.DataTables.Add("PH_PY205B")
    Call oForm.DataSources.DataTables.Add("PH_PY205C")
        
    oGrid01.DataTable = oForm.DataSources.DataTables.Item("PH_PY205A")
    oGrid02.DataTable = oForm.DataSources.DataTables.Item("PH_PY205B")
    oGrid03.DataTable = oForm.DataSources.DataTables.Item("PH_PY205C")
    
    Set oDS_PH_PY205A = oForm.DataSources.DataTables.Item("PH_PY205A")
    Set oDS_PH_PY205B = oForm.DataSources.DataTables.Item("PH_PY205B")
    Set oDS_PH_PY205C = oForm.DataSources.DataTables.Item("PH_PY205C")
    '//----------------------------------------------------------------------------------------------
    '// 기본사항
    '//----------------------------------------------------------------------------------------------
    
    '//////////Folder01//////////_S
    '사업장
    Call oForm.DataSources.UserDataSources.Add("CLTCOD01", dt_SHORT_TEXT, 10)
    Call oForm.Items("CLTCOD01").Specific.DataBind.SetBound(True, "", "CLTCOD01")
    
    '부서
    Call oForm.DataSources.UserDataSources.Add("TeamCode01", dt_SHORT_TEXT, 10)
    Call oForm.Items("TeamCode01").Specific.DataBind.SetBound(True, "", "TeamCode01")
    
    '재직여부
    Call oForm.DataSources.UserDataSources.Add("Status01", dt_SHORT_TEXT, 10)
    Call oForm.Items("Status01").Specific.DataBind.SetBound(True, "", "Status01")
    
    '사원번호
    Call oForm.DataSources.UserDataSources.Add("MSTCOD01", dt_SHORT_TEXT, 20)
    Call oForm.Items("MSTCOD01").Specific.DataBind.SetBound(True, "", "MSTCOD01")
    
    '사원성명
    Call oForm.DataSources.UserDataSources.Add("MSTNAM01", dt_SHORT_TEXT, 50)
    Call oForm.Items("MSTNAM01").Specific.DataBind.SetBound(True, "", "MSTNAM01")

    '기준년도
    Call oForm.DataSources.UserDataSources.Add("StdYear01", dt_SHORT_TEXT, 4)
    Call oForm.Items("StdYear01").Specific.DataBind.SetBound(True, "", "StdYear01")
    
    '기준월
    Call oForm.DataSources.UserDataSources.Add("StdMt01", dt_SHORT_TEXT, 4)
    Call oForm.Items("StdMt01").Specific.DataBind.SetBound(True, "", "StdMt01")
    
    '계획구분
    Call oForm.DataSources.UserDataSources.Add("PlnCls01", dt_SHORT_TEXT, 10)
    Call oForm.Items("PlnCls01").Specific.DataBind.SetBound(True, "", "PlnCls01")
    
    '상태
    Call oForm.DataSources.UserDataSources.Add("DocSts01", dt_SHORT_TEXT, 10)
    Call oForm.Items("DocSts01").Specific.DataBind.SetBound(True, "", "DocSts01")
    
    '교육명
    Call oForm.DataSources.UserDataSources.Add("EduName01", dt_SHORT_TEXT, 200)
    Call oForm.Items("EduName01").Specific.DataBind.SetBound(True, "", "EduName01")
    
    '교육기관
    Call oForm.DataSources.UserDataSources.Add("EduOrg01", dt_SHORT_TEXT, 100)
    Call oForm.Items("EduOrg01").Specific.DataBind.SetBound(True, "", "EduOrg01")
    
    '취소제외
    Call oForm.DataSources.UserDataSources.Add("CancelYN01", dt_SHORT_TEXT, 5)
    Call oForm.Items("CancelYN01").Specific.DataBind.SetBound(True, "", "CancelYN01")
    '//////////Folder01//////////_E
    
    '//////////Folder02//////////_S
    '사업장
    Call oForm.DataSources.UserDataSources.Add("CLTCOD02", dt_SHORT_TEXT, 10)
    Call oForm.Items("CLTCOD02").Specific.DataBind.SetBound(True, "", "CLTCOD02")
    
    '기준년월(시작)
    Call oForm.DataSources.UserDataSources.Add("FrMt02", dt_SHORT_TEXT, 6)
    Call oForm.Items("FrMt02").Specific.DataBind.SetBound(True, "", "FrMt02")
    
    '기준년월(종료)
    Call oForm.DataSources.UserDataSources.Add("ToMt02", dt_SHORT_TEXT, 6)
    Call oForm.Items("ToMt02").Specific.DataBind.SetBound(True, "", "ToMt02")
    
    '재직여부
    Call oForm.DataSources.UserDataSources.Add("Status02", dt_SHORT_TEXT, 10)
    Call oForm.Items("Status02").Specific.DataBind.SetBound(True, "", "Status02")
    
    '수정계획 미출력
    Call oForm.DataSources.UserDataSources.Add("Check201", dt_SHORT_TEXT, 10)
    Call oForm.Items("Check201").Specific.DataBind.SetBound(True, "", "Check201")
    '//////////Folder02//////////_E
    
    '//////////Folder03//////////_S
    '사업장
    Call oForm.DataSources.UserDataSources.Add("CLTCOD03", dt_SHORT_TEXT, 10)
    Call oForm.Items("CLTCOD03").Specific.DataBind.SetBound(True, "", "CLTCOD03")
    
    '부서
    Call oForm.DataSources.UserDataSources.Add("TeamCode03", dt_SHORT_TEXT, 10)
    Call oForm.Items("TeamCode03").Specific.DataBind.SetBound(True, "", "TeamCode03")
    
    '재직여부
    Call oForm.DataSources.UserDataSources.Add("Status03", dt_SHORT_TEXT, 10)
    Call oForm.Items("Status03").Specific.DataBind.SetBound(True, "", "Status03")
    
    '사원번호
    Call oForm.DataSources.UserDataSources.Add("MSTCOD03", dt_SHORT_TEXT, 20)
    Call oForm.Items("MSTCOD03").Specific.DataBind.SetBound(True, "", "MSTCOD03")
    
    '사원성명
    Call oForm.DataSources.UserDataSources.Add("MSTNAM03", dt_SHORT_TEXT, 50)
    Call oForm.Items("MSTNAM03").Specific.DataBind.SetBound(True, "", "MSTNAM03")

    '기준년도
    Call oForm.DataSources.UserDataSources.Add("StdYear03", dt_SHORT_TEXT, 4)
    Call oForm.Items("StdYear03").Specific.DataBind.SetBound(True, "", "StdYear03")
    
    '기준월
    Call oForm.DataSources.UserDataSources.Add("StdMt03", dt_SHORT_TEXT, 4)
    Call oForm.Items("StdMt03").Specific.DataBind.SetBound(True, "", "StdMt03")
    
    '계획구분
    Call oForm.DataSources.UserDataSources.Add("PlnCls03", dt_SHORT_TEXT, 10)
    Call oForm.Items("PlnCls03").Specific.DataBind.SetBound(True, "", "PlnCls03")
    
    '상태
    Call oForm.DataSources.UserDataSources.Add("DocSts03", dt_SHORT_TEXT, 10)
    Call oForm.Items("DocSts03").Specific.DataBind.SetBound(True, "", "DocSts03")
    
    '교육명
    Call oForm.DataSources.UserDataSources.Add("EduName03", dt_SHORT_TEXT, 200)
    Call oForm.Items("EduName03").Specific.DataBind.SetBound(True, "", "EduName03")
    
    '교육기관
    Call oForm.DataSources.UserDataSources.Add("EduOrg03", dt_SHORT_TEXT, 100)
    Call oForm.Items("EduOrg03").Specific.DataBind.SetBound(True, "", "EduOrg03")
    
    '취소제외
    Call oForm.DataSources.UserDataSources.Add("CancelYN03", dt_SHORT_TEXT, 5)
    Call oForm.Items("CancelYN03").Specific.DataBind.SetBound(True, "", "CancelYN03")
    '//////////Folder03//////////_E
    
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
    
PH_PY205_CreateItems_Error:

    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY205_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    
End Function

Sub PH_PY205_ComboBox_Setting()
On Error GoTo PH_PY205_ComboBox_Setting_Error

    Dim oCombo As SAPbouiCOM.ComboBox
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Call oForm.Freeze(True)
    
    '//////////Folder01//////////_S
    '재직여부
    Call oForm.Items("Status01").Specific.ValidValues.Add("1", "재직자")
    Call oForm.Items("Status01").Specific.ValidValues.Add("2", "퇴직자포함")
    Call oForm.Items("Status01").Specific.Select(0, psk_Index)
    
    '계획구분
    sQry = "        SELECT      U_Code AS [Code],"
    sQry = sQry & "             U_CodeNm AS [Name]"
    sQry = sQry & " FROM        [@PS_HR200L]"
    sQry = sQry & " WHERE       Code = 'P239'"
    sQry = sQry & "             AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY    U_Seq"
    
    Call oForm.Items("PlnCls01").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm.Items("PlnCls01").Specific, sQry, "", False, False)
    Call oForm.Items("PlnCls01").Specific.Select(0, psk_Index)
    
    '기준월
    sQry = "        SELECT      U_Code AS [Code],"
    sQry = sQry & "             U_CodeNm AS [Name]"
    sQry = sQry & " FROM        [@PS_HR200L]"
    sQry = sQry & " WHERE       Code = '4'"
    sQry = sQry & "             AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY    U_Seq"
    
    Call oForm.Items("StdMt01").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm.Items("StdMt01").Specific, sQry, "", False, False)
    Call oForm.Items("StdMt01").Specific.Select(0, psk_Index)
    
    '상태
    Call oForm.Items("DocSts01").Specific.ValidValues.Add("%", "전체")
    Call oForm.Items("DocSts01").Specific.ValidValues.Add("O", "계획")
    Call oForm.Items("DocSts01").Specific.ValidValues.Add("P", "완료")
    Call oForm.Items("DocSts01").Specific.ValidValues.Add("C", "취소")
    Call oForm.Items("DocSts01").Specific.Select(0, psk_Index)
    '//////////Folder01//////////_E
    
    '//////////Folder02//////////_S
    '재직여부
    Call oForm.Items("Status02").Specific.ValidValues.Add("1", "재직자")
    Call oForm.Items("Status02").Specific.ValidValues.Add("2", "퇴직자포함")
    Call oForm.Items("Status02").Specific.Select(0, psk_Index)
    '//////////Folder02//////////_E

    '//////////Folder03//////////_S
    '재직여부
    Call oForm.Items("Status03").Specific.ValidValues.Add("1", "재직자")
    Call oForm.Items("Status03").Specific.ValidValues.Add("2", "퇴직자포함")
    Call oForm.Items("Status03").Specific.Select(0, psk_Index)
    
    '계획구분
    sQry = "        SELECT      U_Code AS [Code],"
    sQry = sQry & "             U_CodeNm AS [Name]"
    sQry = sQry & " FROM        [@PS_HR200L]"
    sQry = sQry & " WHERE       Code = 'P239'"
    sQry = sQry & "             AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY    U_Seq"

    Call oForm.Items("PlnCls03").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm.Items("PlnCls03").Specific, sQry, "", False, False)
    Call oForm.Items("PlnCls03").Specific.Select(0, psk_Index)
    
    '월
    sQry = "        SELECT      U_Code AS [Code],"
    sQry = sQry & "             U_CodeNm AS [Name]"
    sQry = sQry & " FROM        [@PS_HR200L]"
    sQry = sQry & " WHERE       Code = '4'"
    sQry = sQry & "             AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY    U_Seq"
    
    Call oForm.Items("StdMt03").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm.Items("StdMt03").Specific, sQry, "", False, False)
    Call oForm.Items("StdMt03").Specific.Select(0, psk_Index)
    
    '상태
    Call oForm.Items("DocSts03").Specific.ValidValues.Add("%", "전체")
    Call oForm.Items("DocSts03").Specific.ValidValues.Add("O", "계획")
    Call oForm.Items("DocSts03").Specific.ValidValues.Add("P", "완료")
    Call oForm.Items("DocSts03").Specific.ValidValues.Add("C", "취소")
    Call oForm.Items("DocSts03").Specific.Select(0, psk_Index)
    '//////////Folder03//////////_E

    '//////////매트릭스//////////_S
'    '상태
'    Call oMat01.Columns("DocStatus").ValidValues.Add("O", "계획")
'    Call oMat01.Columns("DocStatus").ValidValues.Add("P", "완료")
'    Call oMat01.Columns("DocStatus").ValidValues.Add("C", "취소")
'
'    '계획구분
'    sQry = "            SELECT      U_Code AS [Code],"
'    sQry = sQry & "                 U_CodeNm AS [Name]"
'    sQry = sQry & " FROM        [@PS_HR200L]"
'    sQry = sQry & " WHERE       Code = 'P239'"
'    sQry = sQry & "                 AND U_UseYN = 'Y'"
'    sQry = sQry & " ORDER BY  U_Seq"
'    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("PlnCls"), sQry)
'
'    '이수(수료)증
'    Call oMat01.Columns("Certi").ValidValues.Add("O", "미제출")
'    Call oMat01.Columns("Certi").ValidValues.Add("C", "제출")
'
'    '보고서
'    Call oMat01.Columns("Report").ValidValues.Add("O", "미제출")
'    Call oMat01.Columns("Report").ValidValues.Add("C", "제출")
    '//////////매트릭스//////////_E
    
    Call oForm.Freeze(False)
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    
    Exit Sub
PH_PY205_ComboBox_Setting_Error:
    Call oForm.Freeze(False)
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    Sbo_Application.SetStatusBarMessage "PH_PY205_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY205_EnableMenus()
On Error GoTo PH_PY205_EnableMenus_Error

    Call oForm.EnableMenu("1283", True)     '//제거
    Call oForm.EnableMenu("1284", False)    '//취소
    Call oForm.EnableMenu("1293", False)     '//행삭제
    
    Exit Sub
    
PH_PY205_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY205_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY205_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PH_PY205_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY205_FormItemEnabled
    Else
'        oForm.Mode = fm_FIND_MODE
        Call PH_PY205_FormItemEnabled
'        oForm.Items("Code").Specific.VALUE = oFromDocEntry01
'        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY205_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY205_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY205_FormItemEnabled()
On Error GoTo PH_PY205_FormItemEnabled_Error

    Dim oCombo As SAPbouiCOM.ComboBox
    
    Call oForm.Freeze(True)
    If (oForm.Mode = fm_ADD_MODE) Then
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD01")
        Call CLTCOD_Select(oForm, "CLTCOD02")
        Call CLTCOD_Select(oForm, "CLTCOD03")
        
        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", False)    '//문서추가
        
    ElseIf (oForm.Mode = fm_FIND_MODE) Then
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD01")
        Call CLTCOD_Select(oForm, "CLTCOD02")
        Call CLTCOD_Select(oForm, "CLTCOD03")
        
        Call oForm.EnableMenu("1281", False)    '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가
    ElseIf (oForm.Mode = fm_OK_MODE) Then
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD01", False)
        Call CLTCOD_Select(oForm, "CLTCOD02", False)
        Call CLTCOD_Select(oForm, "CLTCOD03", False)
        
        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가

    End If
    Call oForm.Freeze(False)
    Exit Sub
    
PH_PY205_FormItemEnabled_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY205_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY205_FormResize()
On Error GoTo PH_PY205_FormResize_Error
    
    '그룹박스 크기 동적 할당
    oForm.Items("GrpBox01").Height = oForm.Items("Grid01").Height + 145
    oForm.Items("GrpBox01").Width = oForm.Items("Grid01").Width + 30
    
    If oGrid01.Columns.Count > 0 Then
        Call oGrid01.AutoResizeColumns
    End If
    
    If oGrid02.Columns.Count > 0 Then
        oGrid02.AutoResizeColumns
    End If

    If oGrid03.Columns.Count > 0 Then
        oGrid03.AutoResizeColumns
    End If

'    If oGrid04.Columns.Count > 0 Then
'        oGrid04.AutoResizeColumns
'    End If
'
'    If oGrid05.Columns.Count > 0 Then
'        oGrid05.AutoResizeColumns
'    End If
        
    Exit Sub
PH_PY205_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY205_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY205_FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo PH_PY205_FlushToItemValue_Error

    Dim loopCount As Integer
    Dim ErrNum    As Integer
    Dim sQry      As String
    Dim ItemCode  As String
    
    Dim CLTCOD As String

    Dim oCombo As SAPbouiCOM.ComboBox
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = oCompany.GetBusinessObject(BoRecordset)
    
    Select Case oUID

        Case "CLTCOD01"
            
            CLTCOD = Trim(oForm.Items("CLTCOD01").Specific.VALUE)
            
            If oForm.Items("TeamCode01").Specific.ValidValues.Count > 0 Then
                For loopCount = oForm.Items("TeamCode01").Specific.ValidValues.Count - 1 To 0 Step -1
                    Call oForm.Items("TeamCode01").Specific.ValidValues.Remove(loopCount, psk_Index)
                Next
            End If
            
            '부서콤보세팅
            Call oForm.Items("TeamCode01").Specific.ValidValues.Add("%", "전체")
            sQry = "        SELECT      U_Code AS [Code],"
            sQry = sQry & "             U_CodeNm As [Name]"
            sQry = sQry & " FROM        [@PS_HR200L]"
            sQry = sQry & " WHERE       Code = '1'"
            sQry = sQry & "             AND U_UseYN = 'Y'"
            sQry = sQry & "             AND U_Char2 = '" & CLTCOD & "'"
            sQry = sQry & " UNION ALL "
            sQry = sQry & " SELECT      '9999' AS [Code],"
            sQry = sQry & "             '전부서' As [Name]"
            sQry = sQry & " ORDER BY    U_Code"
            Call MDC_SetMod.Set_ComboList(oForm.Items("TeamCode01").Specific, sQry, "", False, False)
            Call oForm.Items("TeamCode01").Specific.Select(0, psk_Index)

        Case "MSTCOD01"

            oForm.Items("MSTNAM01").Specific.VALUE = MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & oForm.Items("MSTCOD01").Specific.VALUE & "'") '성명
        
        Case "CLTCOD03"
            
            CLTCOD = Trim(oForm.Items("CLTCOD03").Specific.VALUE)
            
            If oForm.Items("TeamCode03").Specific.ValidValues.Count > 0 Then
                For loopCount = oForm.Items("TeamCode03").Specific.ValidValues.Count - 1 To 0 Step -1
                    Call oForm.Items("TeamCode03").Specific.ValidValues.Remove(loopCount, psk_Index)
                Next
            End If
            
            '부서콤보세팅
            Call oForm.Items("TeamCode03").Specific.ValidValues.Add("%", "전체")
            sQry = "        SELECT      U_Code AS [Code],"
            sQry = sQry & "             U_CodeNm As [Name]"
            sQry = sQry & " FROM        [@PS_HR200L]"
            sQry = sQry & " WHERE       Code = '1'"
            sQry = sQry & "             AND U_UseYN = 'Y'"
            sQry = sQry & "             AND U_Char2 = '" & CLTCOD & "'"
            sQry = sQry & " ORDER BY    U_Seq"
            Call MDC_SetMod.Set_ComboList(oForm.Items("TeamCode03").Specific, sQry, "", False, False)
            Call oForm.Items("TeamCode03").Specific.Select(0, psk_Index)

        Case "MSTCOD03"

            oForm.Items("MSTNAM03").Specific.VALUE = MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & oForm.Items("MSTCOD03").Specific.VALUE & "'") '성명
            
    End Select

    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    
    Exit Sub
PH_PY205_FlushToItemValue_Error:

    Set oCombo = Nothing
    Set oRecordSet01 = Nothing

    Call MDC_Com.MDC_GF_Message("PH_PY205_FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E")

End Sub

Public Sub Raise_FormItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_FormItemEvent_Error

    Dim sQry    As String
    Dim i       As Long
    Dim oCombo  As SAPbouiCOM.ComboBox
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Select Case pval.EventType
        Case et_ITEM_PRESSED: '//1
            
            If pval.BeforeAction = True Then
            
                If pval.ItemUID = "BtnSrch01" Then
                    
                    Call PH_PY205_DataFind01
                    
                End If
                
                If pval.ItemUID = "BtnSrch02" Then
                    
                    Call PH_PY205_DataFind02
                    
                End If
                
                If pval.ItemUID = "BtnSrch03" Then
                    
                    Call PH_PY205_DataFind03
                    
                End If
                
                If pval.ItemUID = "BtnPrt01" Then
                
                    Call PH_PY205_Print_Report01
                    
                End If
                
                If pval.ItemUID = "BtnPrt02" Then
                
                    Call PH_PY205_Print_Report02
                    
                End If
                
                If pval.ItemUID = "BtnPrt03" Then
                
                    Call PH_PY205_Print_Report03
                    
                End If
            
            ElseIf pval.BeforeAction = False Then
            
                '폴더를 사용할 때는 필수 소스_S
                'Folder01이 선택되었을 때
                If pval.ItemUID = "Folder01" Then
                    
                    oForm.PaneLevel = 1
                    oForm.DefButton = "BtnSrch01"
                    
                    Call PH_PY205_FolderClear
                    
                End If
                
                'Folder02가 선택되었을 때
                If pval.ItemUID = "Folder02" Then
                    
                    oForm.PaneLevel = 2
                    oForm.DefButton = "BtnSrch02"
                    
                    Call PH_PY205_FolderClear
                    
                End If
                
                'Folder03가 선택되었을 때
                If pval.ItemUID = "Folder03" Then
                    
                    oForm.PaneLevel = 3
                    oForm.DefButton = "BtnSrch03"
                    
                    Call PH_PY205_FolderClear
'                    Call PH_PY205_FlushToItemValue("CLTCOD03") '사업장 콤보박스 이벤트 강제 실행(하위 팀 콤보박스의 Binding 재설정을 위함)
                    
                End If
                '폴더를 사용할 때는 필수 소스_E
            
'                If oForm.Mode = fm_ADD_MODE Then
'                    If pval.ActionSuccess = True Then
'                        Call PH_PY205_FormItemEnabled
'                    End If
'                ElseIf oForm.Mode = fm_UPDATE_MODE Then
'                    If pval.ActionSuccess = True Then
'                        Call PH_PY205_FormItemEnabled
'                    End If
'                ElseIf oForm.Mode = fm_OK_MODE Then
'                    If pval.ActionSuccess = True Then
'                        Call PH_PY205_FormItemEnabled
'                    End If
'                End If

            End If
'----------------------------------------------------------
        Case et_KEY_DOWN: '//2

            If pval.BeforeAction = True Then
            
                Call MDC_PS_Common.ActiveUserDefineValue(oForm, pval, BubbleEvent, "MSTCOD01", "") 'Header(Folder01)-사번
                Call MDC_PS_Common.ActiveUserDefineValue(oForm, pval, BubbleEvent, "MSTCOD03", "") 'Header(Folder02)-사번
            
            ElseIf pval.BeforeAction = False Then
            
            End If

'----------------------------------------------------------
        Case et_GOT_FOCUS: '//3
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
'----------------------------------------------------------
        Case et_LOST_FOCUS: '//4

'----------------------------------------------------------
        Case et_COMBO_SELECT: '//5
            Call oForm.Freeze(True)
            If pval.BeforeAction = True Then
                
            ElseIf pval.BeforeAction = False Then
                If pval.ItemChanged = True Then
                
                    Call PH_PY205_FlushToItemValue(pval.ItemUID)
                
                End If
            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_CLICK: '//6
            If pval.BeforeAction = True Then
                Select Case pval.ItemUID
                Case "Grid01"
                    If pval.Row > 0 Then
'                        Call oGrid01.SelectRow(pval.Row, True, False)
                    End If
                End Select
                
                Select Case pval.ItemUID
                Case "Grid01"
                    If pval.Row > 0 Then
                        oLastItemUID = pval.ItemUID
                        oLastColUID = pval.ColUID
                        oLastColRow = pval.Row
                    End If
                Case "Grid02"
                    If pval.Row > 0 Then
                        oLastItemUID = pval.ItemUID
                        oLastColUID = pval.ColUID
                        oLastColRow = pval.Row
                    End If
                Case "Grid03"
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
'----------------------------------------------------------
        Case et_DOUBLE_CLICK: '//7

'----------------------------------------------------------
        Case et_MATRIX_LINK_PRESSED '//8

'----------------------------------------------------------
        Case et_MATRIX_COLLAPSE_PRESSED '//9
'----------------------------------------------------------
        Case et_VALIDATE: '//10
            Call oForm.Freeze(True)
            If pval.BeforeAction = True Then
            
                If pval.ItemChanged = True Then
                
                    Call PH_PY205_FlushToItemValue(pval.ItemUID)
                
                End If
       
            ElseIf pval.BeforeAction = False Then
                If pval.ItemChanged = True Then
                
                End If
            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_MATRIX_LOAD: '//11
            If pval.BeforeAction = True Then
            ElseIf pval.BeforeAction = False Then
                Call PH_PY205_FormItemEnabled
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
                Set oDS_PH_PY205A = Nothing
                Set oGrid01 = Nothing
                
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
                
                Call PH_PY205_FormResize
                
            End If
'----------------------------------------------------------
        Case et_FORM_KEY_DOWN '//22
            
'----------------------------------------------------------
        Case et_FORM_MENU_HILIGHT '//23
            
'----------------------------------------------------------
        Case et_CHOOSE_FROM_LIST '//27
            If pval.BeforeAction = True Then
            ElseIf pval.Before_Action = False Then
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
    Sbo_Application.SetStatusBarMessage "Raise_ItemEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub Raise_FormMenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_FormMenuEvent_Error

    Dim i As Long
    
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
                Call PH_PY205_FormItemEnabled

            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//문서찾기
                Call PH_PY205_FormItemEnabled
            Case "1282":    '//문서추가
                Call PH_PY205_FormItemEnabled

            Case "1288", "1289", "1290", "1291":
                Call PH_PY205_FormItemEnabled
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
    
    Exit Sub
    
Raise_RightClickEvent_Error:

    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    
End Sub

Sub PH_PY205_FormClear()
On Error GoTo PH_PY205_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY205'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PH_PY205_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY205_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY205_FolderClear()
On Error GoTo PH_PY205_FolderClear_Error

    If oForm.PaneLevel = 1 Then
    
        If oGrid01.Columns.Count = 0 Then
        
            Call oForm.Items("MSTCOD01").CLICK(ct_Regular)
        
        End If
    
    ElseIf oForm.PaneLevel = 2 Then
    
        If oGrid02.Columns.Count = 0 Then
        
            Call oForm.Items("FrMt02").CLICK(ct_Regular)
        
        End If
    
    ElseIf oForm.PaneLevel = 3 Then
    
        If oGrid03.Columns.Count = 0 Then
        
            Call oForm.Freeze(True)
            Call PH_PY205_FlushToItemValue("CLTCOD03") '사업장 콤보박스 이벤트 강제 실행(하위 팀 콤보박스의 Binding 재설정을 위함)
            Call oForm.Items("MSTCOD03").CLICK(ct_Regular)
            Call oForm.Freeze(False)
            
        End If
    
    End If

    Exit Sub
    
PH_PY205_FolderClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY205_FolderClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY205_DataValidCheck() As Boolean
On Error GoTo PH_PY205_DataValidCheck_Error

    PH_PY205_DataValidCheck = False
    Dim i       As Long
    Dim sQry    As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    If oForm.Items("CLTCOD01").Specific.VALUE = "" Then
        Sbo_Application.SetStatusBarMessage "사업장은 필수입니다.", bmt_Short, True
        oForm.Items("CLTCOD01").CLICK ct_Regular
        PH_PY205_DataValidCheck = False
        Exit Function
    End If
    
    PH_PY205_DataValidCheck = True
        
    Exit Function
    
    Set oRecordSet = Nothing
    
PH_PY205_DataValidCheck_Error:
    Set oRecordSet = Nothing
    PH_PY205_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY205_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Function PH_PY205_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY205_Validate_Error

    PH_PY205_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY205A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", bmt_Short, True
        PH_PY205_Validate = False
        GoTo PH_PY205_Validate_Exit
    End If
'
    If ValidateType = "수정" Then

    ElseIf ValidateType = "행삭제" Then

    ElseIf ValidateType = "취소" Then

    End If
    
    Set oRecordSet = Nothing
    
    Exit Function
    
PH_PY205_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY205_Validate_Error:
    PH_PY205_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY205_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY205_DataFind01()
'******************************************************************************
'Function ID : PH_PY205_DataFind01()
'해당모듈    : PH_PY205
'기능        : 교육계획상태(개인별상세) 조회
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PH_PY205_DataFind01_Error

    Dim i       As Long
    Dim sQry    As String

    Dim oRecordSet  As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Dim CLTCOD    As String '사업장
    Dim TeamCode  As String '부서
    Dim MSTCOD    As String '사번
    Dim StdYear   As String '기준년도
    Dim StdMt     As String '기준월
    Dim Status    As String '재직여부
    Dim DocStatus As String '상태
    Dim PlnCls    As String '계획구분
    Dim EduName   As String '교육명
    Dim EduOrg    As String '교육기관
    Dim CancelYN  As String '취소제외
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
    
    CLTCOD = Trim(oForm.Items("CLTCOD01").Specific.VALUE) '사업장
    TeamCode = Trim(oForm.Items("TeamCode01").Specific.VALUE) '부서
    MSTCOD = Trim(oForm.Items("MSTCOD01").Specific.VALUE) '사번
    StdYear = Trim(oForm.Items("StdYear01").Specific.VALUE) '기준년도
    StdMt = Trim(oForm.Items("StdMt01").Specific.VALUE) '기준월
    Status = Trim(oForm.Items("Status01").Specific.VALUE) '재직여부
    DocStatus = Trim(oForm.Items("DocSts01").Specific.VALUE) '상태
    PlnCls = Trim(oForm.Items("PlnCls01").Specific.VALUE) '계획구분
    EduName = Trim(oForm.Items("EduName01").Specific.VALUE) '교육명
    EduOrg = Trim(oForm.Items("EduOrg01").Specific.VALUE) '교육기관
    
    If oForm.DataSources.UserDataSources.Item("CancelYN01").VALUE = "Y" Then '취소제외 여부
        CancelYN = "Y"
    Else
        CancelYN = "N"
    End If
    
    sQry = "      EXEC [PH_PY205_01] '"
    sQry = sQry & CLTCOD & "','"  '사업장
    sQry = sQry & TeamCode & "','"  '부서
    sQry = sQry & MSTCOD & "','"  '사번
    sQry = sQry & StdYear & "','"  '기준년도
    sQry = sQry & StdMt & "','"  '기준월
    sQry = sQry & Status & "','" '재직여부
    sQry = sQry & DocStatus & "','" '상태
    sQry = sQry & PlnCls & "','" '계획구분
    sQry = sQry & EduName & "','" '교육명
    sQry = sQry & EduOrg & "','" '교육기관
    sQry = sQry & CancelYN & "'" '취소제외

    Call oForm.Freeze(True)

    Call oDS_PH_PY205A.ExecuteQuery(sQry)

    oGrid01.Columns(12).RightJustified = True
    oGrid01.Columns(13).RightJustified = True
'    oGrid01.Columns(14).RightJustified = True
'    oGrid01.Columns(15).RightJustified = True
    
    Call oForm.Freeze(False)
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set oRecordSet = Nothing
    Exit Sub
    
PH_PY205_DataFind01_Error:

    Call oForm.Freeze(False)

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "PH_PY205_DataFind01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY205_DataFind02()
'******************************************************************************
'Function ID : PH_PY205_DataFind02()
'해당모듈    : PH_PY205
'기능        : 계획VS실적(팀별집계) 조회
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PH_PY205_DataFind02_Error

    Dim i       As Long
    Dim sQry    As String

    Dim oRecordSet  As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Dim CLTCOD  As String '사업장
    Dim FrMt    As String '기준년월(시작)
    Dim ToMt    As String '기준년월(종료)
    Dim Status  As String '재직여부
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
    
    CLTCOD = Trim(oForm.Items("CLTCOD02").Specific.VALUE) '사업장
    FrMt = Trim(oForm.Items("FrMt02").Specific.VALUE) '기준년월(시작)
    ToMt = Trim(oForm.Items("ToMt02").Specific.VALUE) '기준년월(종료)
    Status = Trim(oForm.Items("Status02").Specific.VALUE) '재직여부
    
    sQry = "      EXEC [PH_PY205_03] '"
    sQry = sQry & CLTCOD & "','"  '사업장
    sQry = sQry & FrMt & "','" '기준년월(시작)
    sQry = sQry & ToMt & "','" '기준년월(종료)
    sQry = sQry & Status & "'" '재직여부

    Call oForm.Freeze(True)

    Call oDS_PH_PY205B.ExecuteQuery(sQry)

    oGrid02.Columns(2).RightJustified = True
    oGrid02.Columns(3).RightJustified = True
    oGrid02.Columns(4).RightJustified = True
    oGrid02.Columns(5).RightJustified = True
    oGrid02.Columns(6).RightJustified = True
    oGrid02.Columns(7).RightJustified = True
    oGrid02.Columns(8).RightJustified = True
    oGrid02.Columns(9).RightJustified = True
    oGrid02.Columns(10).RightJustified = True
    
    Call oForm.Freeze(False)
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set oRecordSet = Nothing
    Exit Sub
    
PH_PY205_DataFind02_Error:

    Call oForm.Freeze(False)

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "PH_PY205_DataFind02_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY205_DataFind03()
'******************************************************************************
'Function ID : PH_PY205_DataFind03()
'해당모듈    : PH_PY205
'기능        : 계획VS실적(교육별, 개인별) 조회
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PH_PY205_DataFind03_Error

    Dim i       As Long
    Dim sQry    As String

    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Dim CLTCOD    As String '사업장
    Dim TeamCode  As String '부서
    Dim MSTCOD    As String '사번
    Dim StdYear   As String '기준년도
    Dim StdMt     As String '기준월
    Dim Status    As String '재직여부
    Dim DocStatus As String '상태
    Dim PlnCls    As String '계획구분
    Dim EduName   As String '교육명
    Dim EduOrg    As String '교육기관
    Dim CancelYN  As String '취소제외
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
    
    CLTCOD = Trim(oForm.Items("CLTCOD03").Specific.VALUE) '사업장
    TeamCode = Trim(oForm.Items("TeamCode03").Specific.VALUE) '부서
    MSTCOD = Trim(oForm.Items("MSTCOD03").Specific.VALUE) '사번
    StdYear = Trim(oForm.Items("StdYear03").Specific.VALUE) '기준년도
    StdMt = Trim(oForm.Items("StdMt03").Specific.VALUE) '기준월
    Status = Trim(oForm.Items("Status03").Specific.VALUE) '재직여부
    DocStatus = Trim(oForm.Items("DocSts03").Specific.VALUE) '상태
    PlnCls = Trim(oForm.Items("PlnCls03").Specific.VALUE) '계획구분
    EduName = Trim(oForm.Items("EduName03").Specific.VALUE) '교육명
    EduOrg = Trim(oForm.Items("EduOrg03").Specific.VALUE) '교육기관
    
    If oForm.DataSources.UserDataSources.Item("CancelYN03").VALUE = "Y" Then '취소제외 여부
        CancelYN = "Y"
    Else
        CancelYN = "N"
    End If
    
    sQry = "      EXEC [PH_PY205_05] '"
    sQry = sQry & CLTCOD & "','"  '사업장
    sQry = sQry & TeamCode & "','"  '부서
    sQry = sQry & MSTCOD & "','"  '사번
    sQry = sQry & StdYear & "','"  '기준년도
    sQry = sQry & StdMt & "','" '기준월
    sQry = sQry & Status & "','" '재직여부
    sQry = sQry & DocStatus & "','" '상태
    sQry = sQry & PlnCls & "','" '계획구분
    sQry = sQry & EduName & "','" '교육명
    sQry = sQry & EduOrg & "','" '교육기관
    sQry = sQry & CancelYN & "'" '취소제외

    Call oForm.Freeze(True)

    Call oDS_PH_PY205C.ExecuteQuery(sQry)

    oGrid03.Columns(11).RightJustified = True
    oGrid03.Columns(12).RightJustified = True
    oGrid03.Columns(13).RightJustified = True
    oGrid03.Columns(14).RightJustified = True
    oGrid03.Columns(15).RightJustified = True
    oGrid03.Columns(16).RightJustified = True
    
    Call oForm.Freeze(False)
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set oRecordSet = Nothing
    Exit Sub
    
PH_PY205_DataFind03_Error:

    Call oForm.Freeze(False)
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "PH_PY205_DataFind03_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

'Private Sub PH_PY205_TitleSetting(iRow As Long)
'    Dim i               As Long
'    Dim j               As Long
'    Dim sQry            As String
'
'    Dim COLNAM(6)       As String
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
'    COLNAM(0) = "구분"
'    COLNAM(1) = "부서"
'    COLNAM(2) = "담당"
'    COLNAM(3) = "사번"
'    COLNAM(4) = "성명"
'    COLNAM(5) = "횟수"
'    COLNAM(6) = "교통비"
'
'    For i = 0 To UBound(COLNAM)
'        oGrid01.Columns.Item(i).TitleObject.Caption = COLNAM(i)
'        If i >= 0 And i < UBound(COLNAM) Then
'            oGrid01.Columns.Item(i).Editable = False
''        ElseIf i = UBound(COLNAM) Then
''            oGrid01.Columns.Item(i).Editable = True
''            oGrid01.Columns.Item(i).Type = gct_ComboBox
''            Set oComboCol = oGrid01.Columns.Item("Activity")
''
''            sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
''            sQry = sQry & " WHERE U_Char2 = '" & oForm.Items("CLTCOD").Specific.Value & "' And Code = 'P127' AND U_UseYN= 'Y' Order by U_Seq"
''            oRecordSet.DoQuery sQry
''            If oRecordSet.RecordCount > 0 Then
''                    oComboCol.ValidValues.Add "", ""
''                For j = 0 To oRecordSet.RecordCount - 1
''                    oComboCol.ValidValues.Add oRecordSet.Fields(0).Value, oRecordSet.Fields(1).Value
''                    oRecordSet.MoveNext
''                Next j
'''                oComboCol.Select 0, psk_Index
''            End If
''
''
''            oComboCol.DisplayType = cdt_Description
'        End If
'
'    Next i
'
'    oGrid01.AutoResizeColumns
'
'    oForm.Freeze False
'
'    Set oColumn = Nothing
'
'    Exit Sub

'Error_Message:
'    oForm.Freeze False
'    Set oColumn = Nothing
'    Sbo_Application.SetStatusBarMessage "PH_PY205_TitleSetting Error : " & Space(10) & Err.Description, bmt_Short, True
'End Sub

Private Sub PH_PY205_Print_Report01()
On Error GoTo Print_Query_Error

    Dim DocNum     As String
    Dim WinTitle   As String
    Dim ReportName As String
    Dim sQry       As String

    Dim i          As Integer
    Dim ErrNum     As Integer
    Dim Sub_sQry   As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)

    Call ConnectODBC
    
    Dim CLTCOD    As String '사업장
    Dim TeamCode  As String '부서
    Dim MSTCOD    As String '사번
    Dim StdYear   As String '기준년도
    Dim StdMt     As String '기준월
    Dim Status    As String '재직여부
    Dim DocStatus As String '상태
    Dim PlnCls    As String '계획구분
    Dim EduName   As String '교육명
    Dim EduOrg    As String '교육기관
    Dim CancelYN  As String '취소제외
    
    CLTCOD = Trim(oForm.Items("CLTCOD01").Specific.VALUE) '사업장
    TeamCode = Trim(oForm.Items("TeamCode01").Specific.VALUE) '부서
    MSTCOD = Trim(oForm.Items("MSTCOD01").Specific.VALUE) '사번
    StdYear = Trim(oForm.Items("StdYear01").Specific.VALUE) '기준년도
    StdMt = Trim(oForm.Items("StdMt01").Specific.VALUE) '기준월
    Status = Trim(oForm.Items("Status01").Specific.VALUE) '재직여부
    DocStatus = Trim(oForm.Items("DocSts01").Specific.VALUE) '상태
    PlnCls = Trim(oForm.Items("PlnCls01").Specific.VALUE) '계획구분
    EduName = Trim(oForm.Items("EduName01").Specific.VALUE) '교육명
    EduOrg = Trim(oForm.Items("EduOrg01").Specific.VALUE) '교육기관
    
    If oForm.DataSources.UserDataSources.Item("CancelYN01").VALUE = "Y" Then '취소제외 여부
        CancelYN = "Y"
    Else
        CancelYN = "N"
    End If
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    WinTitle = "[PH_PY205] 레포트"
    
    '쿼리

    ReportName = "PH_PY205_01.rpt"

    sQry = "      EXEC [PH_PY205_02] '"
    sQry = sQry & CLTCOD & "','"  '사업장
    sQry = sQry & TeamCode & "','"  '부서
    sQry = sQry & MSTCOD & "','"  '사번
    sQry = sQry & StdYear & "','"  '기준년도
    sQry = sQry & StdMt & "','"  '기준월
    sQry = sQry & Status & "','" '재직여부
    sQry = sQry & DocStatus & "','" '상태
    sQry = sQry & PlnCls & "','" '계획구분
    sQry = sQry & EduName & "','" '교육명
    sQry = sQry & EduOrg & "','" '교육기관
    sQry = sQry & CancelYN & "'" '취소제외
        
    '// Formula 수식필드
    ReDim gRpt_Formula(1)
    ReDim gRpt_Formula_Value(1)
       
    '// SubReport
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)
    
    gRpt_SFormula(1, 1) = ""
    gRpt_SFormula_Value(1, 1) = ""

'    Call oRecordSet.DoQuery(sQry)
'
'    If oRecordSet.RecordCount = 0 Then
'        ErrNum = 1
'        GoTo Print_Query_Error
'    End If

    '/ Action (sub_query가 있을때는 'Y'로...)/
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V", , 2) = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
    End If
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set oRecordSet = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Print_Query_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set oRecordSet = Nothing
    
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "PH_PY205_Print_Report01_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub

Private Sub PH_PY205_Print_Report02()
On Error GoTo Print_Query_Error

    Dim DocNum     As String
    Dim WinTitle   As String
    Dim ReportName As String
    Dim sQry       As String

    Dim i          As Integer
    Dim ErrNum     As Integer
    Dim Sub_sQry   As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    Dim ProgBar01  As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)

    Call ConnectODBC
    
    Dim CLTCOD As String '사업장
    Dim FrMt   As String '기준년월(시작)
    Dim ToMt   As String '기준년월(종료)
    Dim Status As String '재직여부
    
    CLTCOD = Trim(oForm.Items("CLTCOD02").Specific.VALUE) '사업장
    FrMt = Trim(oForm.Items("FrMt02").Specific.VALUE) '기준년월(시작)
    ToMt = Trim(oForm.Items("ToMt02").Specific.VALUE) '기준년월(종료)
    Status = Trim(oForm.Items("Status02").Specific.VALUE) '재직여부
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    WinTitle = "[PH_PY205] 레포트"
    
    If oForm.DataSources.UserDataSources.Item("Check201").VALUE = "Y" Then '수정계획 미출력
        ReportName = "PH_PY205_04.rpt"
    Else
        ReportName = "PH_PY205_02.rpt"
    End If
    
    

    sQry = "      EXEC [PH_PY205_04] '"
    sQry = sQry & CLTCOD & "','"  '사업장
    sQry = sQry & FrMt & "','" '기준년월(시작)
    sQry = sQry & ToMt & "','" '기준년월(종료)
    sQry = sQry & Status & "'" '재직여부
        
    '// Formula 수식필드
    ReDim gRpt_Formula(1)
    ReDim gRpt_Formula_Value(1)
       
    '// SubReport
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)
    
    gRpt_SFormula(1, 1) = ""
    gRpt_SFormula_Value(1, 1) = ""

'    Call oRecordSet.DoQuery(sQry)
'
'    If oRecordSet.RecordCount = 0 Then
'        ErrNum = 1
'        GoTo Print_Query_Error
'    End If

    '/ Action (sub_query가 있을때는 'Y'로...)/
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V", , 2) = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
    End If
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set oRecordSet = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Print_Query_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set oRecordSet = Nothing
    
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "PH_PY205_Print_Report02_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub

Private Sub PH_PY205_Print_Report03()
On Error GoTo Print_Query_Error

    Dim DocNum      As String
    Dim WinTitle    As String
    Dim ReportName  As String
    Dim sQry        As String

    Dim i           As Integer
    Dim ErrNum      As Integer
    Dim Sub_sQry    As String
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    Dim ProgBar01   As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)

    Call ConnectODBC
    
    Dim CLTCOD    As String '사업장
    Dim TeamCode  As String '부서
    Dim MSTCOD    As String '사번
    Dim StdYear   As String '기준년도
    Dim StdMt     As String '기준월
    Dim Status    As String '재직여부
    Dim DocStatus As String '상태
    Dim PlnCls    As String '계획구분
    Dim EduName   As String '교육명
    Dim EduOrg    As String '교육기관
    Dim CancelYN  As String '취소제외
    
    CLTCOD = Trim(oForm.Items("CLTCOD03").Specific.VALUE) '사업장
    TeamCode = Trim(oForm.Items("TeamCode03").Specific.VALUE) '부서
    MSTCOD = Trim(oForm.Items("MSTCOD03").Specific.VALUE) '사번
    StdYear = Trim(oForm.Items("StdYear03").Specific.VALUE) '기준년도
    StdMt = Trim(oForm.Items("StdMt03").Specific.VALUE) '기준월
    Status = Trim(oForm.Items("Status03").Specific.VALUE) '재직여부
    DocStatus = Trim(oForm.Items("DocSts03").Specific.VALUE) '상태
    PlnCls = Trim(oForm.Items("PlnCls03").Specific.VALUE) '계획구분
    EduName = Trim(oForm.Items("EduName03").Specific.VALUE) '교육명
    EduOrg = Trim(oForm.Items("EduOrg03").Specific.VALUE) '교육기관
    
    If oForm.DataSources.UserDataSources.Item("CancelYN03").VALUE = "Y" Then '취소제외 여부
        CancelYN = "Y"
    Else
        CancelYN = "N"
    End If
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    WinTitle = "[PH_PY205] 레포트"
    
    '쿼리

    ReportName = "PH_PY205_03.rpt"

    sQry = "      EXEC [PH_PY205_06] '"
    sQry = sQry & CLTCOD & "','"  '사업장
    sQry = sQry & TeamCode & "','"  '부서
    sQry = sQry & MSTCOD & "','"  '사번
    sQry = sQry & StdYear & "','"  '기준년도
    sQry = sQry & StdMt & "','"  '기준월
    sQry = sQry & Status & "','" '재직여부
    sQry = sQry & DocStatus & "','" '상태
    sQry = sQry & PlnCls & "','" '계획구분
    sQry = sQry & EduName & "','" '교육명
    sQry = sQry & EduOrg & "','" '교육기관
    sQry = sQry & CancelYN & "'" '취소제외
        
    '// Formula 수식필드
    ReDim gRpt_Formula(1)
    ReDim gRpt_Formula_Value(1)
       
    '// SubReport
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)
    
    gRpt_SFormula(1, 1) = ""
    gRpt_SFormula_Value(1, 1) = ""

'    Call oRecordSet.DoQuery(sQry)
'
'    If oRecordSet.RecordCount = 0 Then
'        ErrNum = 1
'        GoTo Print_Query_Error
'    End If

    '/ Action (sub_query가 있을때는 'Y'로...)/
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V", , 2) = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
    End If
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set oRecordSet = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Print_Query_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set oRecordSet = Nothing
    
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "PH_PY205_Print_Report03_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub
