VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY001.cls
'//  Module         : 인사관리 > 인사관리
'//  Desc           : 사원마스터등록
'//********************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form

Public oMat1 As SAPbouiCOM.Matrix
Public oMat2 As SAPbouiCOM.Matrix
Public oMat3 As SAPbouiCOM.Matrix
Public oMat4 As SAPbouiCOM.Matrix
Public oMat5 As SAPbouiCOM.Matrix
Public oMat6 As SAPbouiCOM.Matrix
Public oMat7 As SAPbouiCOM.Matrix
Public oMat8 As SAPbouiCOM.Matrix
Public oMat9 As SAPbouiCOM.Matrix
Public oMat10 As SAPbouiCOM.Matrix
Public oMat11 As SAPbouiCOM.Matrix
Public oMat12 As SAPbouiCOM.Matrix
Public oMat13 As SAPbouiCOM.Matrix
Public oMat14 As SAPbouiCOM.Matrix
Public oMat15 As SAPbouiCOM.Matrix

Private oDS_PH_PY001A As SAPbouiCOM.DBDataSource
Private oDS_PH_PY001B As SAPbouiCOM.DBDataSource
Private oDS_PH_PY001C As SAPbouiCOM.DBDataSource
Private oDS_PH_PY001D As SAPbouiCOM.DBDataSource
Private oDS_PH_PY001E As SAPbouiCOM.DBDataSource
Private oDS_PH_PY001F As SAPbouiCOM.DBDataSource
Private oDS_PH_PY001G As SAPbouiCOM.DBDataSource
Private oDS_PH_PY001H As SAPbouiCOM.DBDataSource
Private oDS_PH_PY001J As SAPbouiCOM.DBDataSource
Private oDS_PH_PY001K As SAPbouiCOM.DBDataSource
Private oDS_PH_PY001L As SAPbouiCOM.DBDataSource
Private oDS_PH_PY001M As SAPbouiCOM.DBDataSource
Private oDS_PH_PY001N As SAPbouiCOM.DBDataSource
Private oDS_PH_PY001P As SAPbouiCOM.DBDataSource
Private oDS_PH_PY001Q As SAPbouiCOM.DBDataSource
Private oDS_PH_PY001R As SAPbouiCOM.DBDataSource

Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long

Private Const FO_COPY = &H2&
Private UserPic     As String

Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY001.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY001_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY001"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)
    
    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
    oForm.DataBrowser.BrowseBy = "Code"
    oForm.PaneLevel = 1
    
    oForm.Freeze True
    Call PH_PY001_CreateItems
    Call PH_PY001_EnableMenus
    Call PH_PY001_SetDocument(oFromDocEntry01)
'    Call PH_PY001_FormResize
    
    oForm.Update
    oForm.Freeze False
    
    oForm.Items("FLD01").Specific.Select '폼이 로드 될 때 첫번 째 폴더를 선택됨
    
    
    
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

Private Function PH_PY001_CreateItems() As Boolean

    Dim sQry        As String
    Dim i           As Long
    
    Dim oCheck      As SAPbouiCOM.CheckBox
    Dim oEdit       As SAPbouiCOM.EditText
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns    As SAPbouiCOM.Columns
    Dim optBtn      As SAPbouiCOM.OptionBtn
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY001_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Set oDS_PH_PY001A = oForm.DataSources.DBDataSources("@PH_PY001A")
    Set oDS_PH_PY001B = oForm.DataSources.DBDataSources("@PH_PY001B")
    Set oDS_PH_PY001C = oForm.DataSources.DBDataSources("@PH_PY001C")
    Set oDS_PH_PY001D = oForm.DataSources.DBDataSources("@PH_PY001D")
    Set oDS_PH_PY001E = oForm.DataSources.DBDataSources("@PH_PY001E")
    Set oDS_PH_PY001F = oForm.DataSources.DBDataSources("@PH_PY001F")
    Set oDS_PH_PY001G = oForm.DataSources.DBDataSources("@PH_PY001G")
    Set oDS_PH_PY001H = oForm.DataSources.DBDataSources("@PH_PY001H")
    Set oDS_PH_PY001J = oForm.DataSources.DBDataSources("@PH_PY001J")
    Set oDS_PH_PY001K = oForm.DataSources.DBDataSources("@PH_PY001K")
    Set oDS_PH_PY001L = oForm.DataSources.DBDataSources("@PH_PY001L")
    Set oDS_PH_PY001M = oForm.DataSources.DBDataSources("@PH_PY001M")
    Set oDS_PH_PY001N = oForm.DataSources.DBDataSources("@PH_PY001N")
    Set oDS_PH_PY001P = oForm.DataSources.DBDataSources("@PH_PY001P")
    Set oDS_PH_PY001Q = oForm.DataSources.DBDataSources("@PH_PY001Q")
    Set oDS_PH_PY001R = oForm.DataSources.DBDataSources("@PH_PY001R")
    
    Set oMat1 = oForm.Items("Mat1").Specific            '//@PH_PY001B
    Set oMat2 = oForm.Items("Mat2").Specific            '//@PH_PY001C
    Set oMat3 = oForm.Items("Mat3").Specific            '//@PH_PY001D
    Set oMat4 = oForm.Items("Mat4").Specific            '//@PH_PY001E
    Set oMat5 = oForm.Items("Mat5").Specific            '//@PH_PY001F
    Set oMat6 = oForm.Items("Mat6").Specific            '//@PH_PY001G
    Set oMat7 = oForm.Items("Mat7").Specific            '//@PH_PY001H
    Set oMat8 = oForm.Items("Mat8").Specific            '//@PH_PY001J
    Set oMat9 = oForm.Items("Mat9").Specific            '//@PH_PY001K
    Set oMat10 = oForm.Items("Mat10").Specific          '//@PH_PY001L
    Set oMat11 = oForm.Items("Mat11").Specific          '//@PH_PY001M
    Set oMat12 = oForm.Items("Mat12").Specific          '//@PH_PY001N
    Set oMat13 = oForm.Items("Mat13").Specific          '//@PH_PY001P
    Set oMat14 = oForm.Items("Mat14").Specific          '//@PH_PY001Q
    Set oMat15 = oForm.Items("Mat15").Specific          '//@PH_PY001Q
        
    oMat1.SelectionMode = ms_NotSupported
    oMat1.AutoResizeColumns
    oMat2.SelectionMode = ms_NotSupported
    oMat2.AutoResizeColumns
    oMat3.SelectionMode = ms_NotSupported
    oMat3.AutoResizeColumns
    oMat4.SelectionMode = ms_NotSupported
    oMat4.AutoResizeColumns
    oMat5.SelectionMode = ms_NotSupported
    oMat5.AutoResizeColumns
    oMat6.SelectionMode = ms_NotSupported
    oMat6.AutoResizeColumns
    oMat7.SelectionMode = ms_NotSupported
    oMat7.AutoResizeColumns
    oMat8.SelectionMode = ms_NotSupported
    oMat8.AutoResizeColumns
    oMat9.SelectionMode = ms_NotSupported
    oMat9.AutoResizeColumns
    oMat10.SelectionMode = ms_NotSupported
    oMat10.AutoResizeColumns
    oMat11.SelectionMode = ms_NotSupported
    oMat11.AutoResizeColumns
    oMat12.SelectionMode = ms_NotSupported
    oMat12.AutoResizeColumns
    oMat13.SelectionMode = ms_NotSupported
    oMat13.AutoResizeColumns
    oMat14.SelectionMode = ms_NotSupported
    oMat14.AutoResizeColumns
    oMat15.SelectionMode = ms_NotSupported
    oMat15.AutoResizeColumns

    '//----------------------------------------------------------------------------------------------
    '// 기본사항
    '//----------------------------------------------------------------------------------------------
    oForm.AutoManaged = True
    Call AutoManaged(oForm, "Code,FullName,CLTCOD")

    '//사업장
    Set oCombo = oForm.Items("CLTCOD").Specific
'    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
'    Call SetReDataCombo(oForm, sQry, oCombo)
    oForm.Items("CLTCOD").DisplayDesc = True

    '// 성
    Set oCombo = oForm.Items("sex").Specific
    oCombo.ValidValues.Add "M", "남자"
    oCombo.ValidValues.Add "F", "여자"
'    oCombo.Select 0, psk_Index
    oForm.Items("sex").DisplayDesc = True

    '// 음양력
    Set oCombo = oForm.Items("CldType").Specific
    oCombo.ValidValues.Add "S", "양력"
    oCombo.ValidValues.Add "R", "음력"
    oCombo.Select 0, psk_Index
    oForm.Items("CldType").DisplayDesc = True

    '// 최종학력
    Set oCombo = oForm.Items("lastEduc").Specific
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P123' AND U_UseYN= 'Y'"
    Call SetReDataCombo(oForm, sQry, oCombo)
    oForm.Items("lastEduc").DisplayDesc = True

    '//부서
    oForm.Items("TeamCode").DisplayDesc = True

    '//담당
    oForm.Items("RspCode").DisplayDesc = True

    '//반
    oForm.Items("ClsCode").DisplayDesc = True

    '// 재직구분
    Set oCombo = oForm.Items("status").Specific
    sQry = "SELECT statusID,name FROM [OHST] ORDER BY statusID"
    Call SetReDataCombo(oForm, sQry, oCombo)
    oForm.Items("status").DisplayDesc = True

    '// 직책
    Set oCombo = oForm.Items("position").Specific
    sQry = "SELECT posID,name FROM [OHPS] ORDER BY posID"
    Call SetReDataCombo(oForm, sQry, oCombo)
    oForm.Items("position").DisplayDesc = True

    '// 직원구분
    Set oCombo = oForm.Items("JIGTYP").Specific
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P126' AND U_UseYN= 'Y' "
    Call SetReDataCombo(oForm, sQry, oCombo)
    oForm.Items("JIGTYP").DisplayDesc = True

    '// 전문직호칭
    Set oCombo = oForm.Items("CallName").Specific
    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P158' AND U_UseYN= 'Y' "
    Call SetReDataCombo(oForm, sQry, oCombo)
    oForm.Items("CallName").DisplayDesc = True

    '// 혼인여부
    Set oCombo = oForm.Items("martStat").Specific
    oCombo.ValidValues.Add "N", "미혼"
    oCombo.ValidValues.Add "Y", "기혼"
'    oCombo.Select 0, psk_Index
    oForm.Items("martStat").DisplayDesc = True

    '// 국적
    Set oCombo = oForm.Items("brthCntr").Specific
    sQry = "SELECT Code,Name FROM [OCRY] ORDER BY Code"
    Call SetReDataCombo(oForm, sQry, oCombo)
    oForm.Items("brthCntr").DisplayDesc = True

    '// 근무형태
    Set oCombo = oForm.Items("ShiftDat").Specific
    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P154' AND U_UseYN= 'Y' "
    Call SetReDataCombo(oForm, sQry, oCombo)
    oForm.Items("ShiftDat").DisplayDesc = True

    '// 근무조
    Set oCombo = oForm.Items("GNMUJO").Specific
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P155' AND U_UseYN= 'Y'"
    Call SetReDataCombo(oForm, sQry, oCombo)
    oForm.Items("GNMUJO").DisplayDesc = True
    
    '//퇴직급여 계약 구분
    Set oCombo = oForm.Items("DBDCDIV").Specific
    oCombo.ValidValues.Add "10", "DB형"
    oCombo.ValidValues.Add "20", "DC형"
    oCombo.Select 0, psk_Index
    oForm.Items("DBDCDIV").DisplayDesc = True

    '//비상연락처 대상 관계
    Set oCombo = oForm.Items("EmRel").Specific
    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P121' AND U_UseYN= 'Y' ORDER BY CAST(U_Code AS NUMERIC) "
    Call SetReDataCombo(oForm, sQry, oCombo)
    Call oCombo.Select(0, psk_Index)
    oForm.Items("EmRel").DisplayDesc = True

    Set oCheck = oForm.Items("Inform01").Specific
    oCheck.ValOn = "Y": oCheck.ValOff = "N"
    oCheck.Checked = False
    Set oCheck = oForm.Items("Inform02").Specific
    oCheck.ValOn = "Y": oCheck.ValOff = "N"
    oCheck.Checked = False
    Set oCheck = oForm.Items("Inform03").Specific
    oCheck.ValOn = "Y": oCheck.ValOff = "N"
    oCheck.Checked = False
    Set oCheck = oForm.Items("Inform04").Specific
    oCheck.ValOn = "Y": oCheck.ValOff = "N"
    oCheck.Checked = False
    Set oCheck = oForm.Items("Inform05").Specific
    oCheck.ValOn = "Y": oCheck.ValOff = "N"
    oCheck.Checked = False
    Set oCheck = oForm.Items("Inform06").Specific
    oCheck.ValOn = "Y": oCheck.ValOff = "N"
    oCheck.Checked = False
    '//----------------------------------------------------------------------------------------------
    '//급여기본
    '//----------------------------------------------------------------------------------------------

    '// 1.급여형태
    Set oCombo = oForm.Items("PAYTYP").Specific
    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P132' AND U_UseYN= 'Y' ORDER BY CAST(U_Code AS NUMERIC) "
    Call SetReDataCombo(oForm, sQry, oCombo)

    '// 2.직급형태
    Set oCombo = oForm.Items("JIGCOD").Specific
    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P129' AND U_UseYN= 'Y' ORDER BY U_Code "
    Call SetReDataCombo(oForm, sQry, oCombo)

    '// 3.원가구분
    Set oCombo = oForm.Items("ACCCOD").Specific
    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P140' AND U_UseYN= 'Y' ORDER BY CAST(U_Code AS NUMERIC) "
    Call SetReDataCombo(oForm, sQry, oCombo)

    '// 4.내국인구분
    Set oCombo = oForm.Items("INTGBN").Specific
    oDS_PH_PY001A.setValue "U_INTGBN", 0, "1"
    oCombo.ValidValues.Add "1", "내국인"
    oCombo.ValidValues.Add "9", "외국인"

    '// 5.거주자구분
    Set oCombo = oForm.Items("DWEGBN").Specific
    oDS_PH_PY001A.setValue "U_DWEGBN", 0, "1"
    oCombo.ValidValues.Add "1", "거주자"
    oCombo.ValidValues.Add "2", "비거주자"

    '// 6.외국인단일세율적용
    Set oCombo = oForm.Items("FRGTAX").Specific
    oDS_PH_PY001A.setValue "U_FRGTAX", 0, "2"
    oCombo.ValidValues.Add "1", "적용함"
    oCombo.ValidValues.Add "2", "적용안함"

    '// 7.급여지급대상
    Set oCombo = oForm.Items("PAYSEL").Specific
    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P213' AND U_UseYN= 'Y' ORDER BY CAST(U_Code AS NUMERIC) "
    Call SetReDataCombo(oForm, sQry, oCombo)

    '// 8.건강보험해외파견경감율
    Set oCombo = oForm.Items("MEDFRG").Specific
    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P193' AND U_UseYN= 'Y' ORDER BY CAST(U_Code AS NUMERIC) "
    Call SetReDataCombo(oForm, sQry, oCombo)
    '// 9.세대주여부
    Set oCombo = oForm.Items("HUSMAN").Specific
    oCombo.ValidValues.Add "1", "세대주"
    oCombo.ValidValues.Add "2", "세대원"

'    If oForm.Mode <> fm_FIND_MODE And oCombo.ValidValues.Count > 1 Then
'        oCombo.Select 1, psk_Index
'    End If

    '/ 고용보험여부
    Set oCheck = oForm.Items("GBHSEL").Specific
    oCheck.ValOn = "Y": oCheck.ValOff = "N"
    oCheck.Checked = False
    '/ 상여지급여부
    Set oCheck = oForm.Items("BNSSEL").Specific
    oCheck.ValOn = "Y": oCheck.ValOff = "N"

    '/ 세액계산여부
    Set oCheck = oForm.Items("TAXSEL").Specific
    oCheck.ValOn = "Y": oCheck.ValOff = "N"
    oCheck.Checked = False
    '/ 국외비과세여부
    Set oCheck = oForm.Items("FRGSEL").Specific
    oCheck.ValOn = "Y": oCheck.ValOff = "N"
    oCheck.Checked = False
    '/ 생산비과세여부
    Set oCheck = oForm.Items("BX1SEL").Specific
    oCheck.ValOn = "Y": oCheck.ValOff = "N"
    oCheck.Checked = False
    '/ 연말정산신고
    Set oCheck = oForm.Items("JSNSEL").Specific
    oCheck.ValOn = "Y": oCheck.ValOff = "N"
    oCheck.Checked = False
    '/ 배우자유무
    Set oCheck = oForm.Items("BAEWOO").Specific
    oCheck.ValOn = "Y": oCheck.ValOff = "N"
    oCheck.Checked = False
    '/ 부녀자공제대상
    Set oCheck = oForm.Items("MZBURI").Specific
    oCheck.ValOn = "Y": oCheck.ValOff = "N"
    oCheck.Checked = False
    '/ 고용보험 납부대상자
    Set oCheck = oForm.Items("NJCGBN").Specific
    oCheck.ValOn = "Y": oCheck.ValOff = "N"
    oCheck.Checked = False
    '/ 본인장애자유무
    Set oCheck = oForm.Items("BJNGAE").Specific
    oCheck.ValOn = "Y": oCheck.ValOff = "N"
    oCheck.Checked = False
    '/ 국민연금연장신청여부
    Set oCheck = oForm.Items("GBHSEL").Specific
    oCheck.ValOn = "Y": oCheck.ValOff = "N"
    oCheck.Checked = False
    '/은행1
    Set oCombo = oForm.Items("BANK1").Specific
    sQry = "SELECT BankCode, BankName From ODSC"
    oRecordSet.DoQuery sQry

    If oRecordSet.RecordCount > 0 Then
         Do Until oRecordSet.EOF
            oCombo.ValidValues.Add oRecordSet.Fields(0).VALUE, oRecordSet.Fields(1).VALUE
            oRecordSet.MoveNext
         Loop
    End If
    oForm.Items("BANK1").DisplayDesc = True
    '/은행2
    Set oCombo = oForm.Items("BANK2").Specific
    sQry = "SELECT BankCode, BankName From ODSC"
    oRecordSet.DoQuery sQry

    If oRecordSet.RecordCount > 0 Then
         Do Until oRecordSet.EOF
            oCombo.ValidValues.Add oRecordSet.Fields(0).VALUE, oRecordSet.Fields(1).VALUE
            oRecordSet.MoveNext
         Loop
    End If
    oForm.Items("BANK2").DisplayDesc = True
    '//----------------------------------------------------------------------------------------------
    '//급여기본(Mat1)
    '//----------------------------------------------------------------------------------------------
    Set oColumn = oMat1.Columns("FILD01")
    oColumn.Editable = True
    Set oColumn = oMat1.Columns("FILD02")
    oColumn.Editable = True
    '//----------------------------------------------------------------------------------------------
    '//급여기본(Mat2)
    '//----------------------------------------------------------------------------------------------
    Set oColumn = oMat2.Columns("FILD01")
    oColumn.Editable = True
    Set oColumn = oMat2.Columns("FILD02")
    oColumn.Editable = True
    '//----------------------------------------------------------------------------------------------
    '//가족사항(Mat3)
    '//----------------------------------------------------------------------------------------------

    '// 생일구분
    Set oColumn = oMat3.Columns("BIRGBN")
    oColumn.ValidValues.Add "S", "양력"
    oColumn.ValidValues.Add "R", "음력"
    oColumn.DisplayDesc = True

    '// 학력
    Set oColumn = oMat3.Columns("FamSch")
    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P123' AND U_UseYN= 'Y' ORDER BY CAST(U_Code AS NUMERIC) "
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        Do Until oRecordSet.EOF
            oColumn.ValidValues.Add Trim$(oRecordSet.Fields(0).VALUE), Trim$(oRecordSet.Fields(1).VALUE)
            oRecordSet.MoveNext
        Loop
    End If
    oColumn.DisplayDesc = True


    '// 본인과의 관계
    Set oColumn = oMat3.Columns("FamGun")
    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P121' AND U_UseYN= 'Y' ORDER BY CAST(U_Code AS NUMERIC) "
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        Do Until oRecordSet.EOF
            oColumn.ValidValues.Add Trim$(oRecordSet.Fields(0).VALUE), Trim$(oRecordSet.Fields(1).VALUE)
            oRecordSet.MoveNext
        Loop
    End If
    oColumn.DisplayDesc = True

    '// 동거여부
    Set oColumn = oMat3.Columns("JsnType")
    oColumn.ValidValues.Add "N", "N"
    oColumn.ValidValues.Add "Y", "Y"
    oColumn.DisplayDesc = True

    '// 관계코드
    Set oColumn = oMat3.Columns("ChkCod")
    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P159' AND U_UseYN= 'Y' ORDER BY CAST(U_Code AS NUMERIC) "
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        Do Until oRecordSet.EOF
            oColumn.ValidValues.Add Trim$(oRecordSet.Fields(0).VALUE), Trim$(oRecordSet.Fields(1).VALUE)
            oRecordSet.MoveNext
        Loop
    End If
    oColumn.DisplayDesc = True

    '// 내외국인
    Set oColumn = oMat3.Columns("ChkInt")
    oColumn.ValidValues.Add "1", "내국인"
    oColumn.ValidValues.Add "9", "외국인"
    oColumn.DisplayDesc = True

    '//----------------------------------------------------------------------------------------------
    '//경력사항(Mat4)
    '//----------------------------------------------------------------------------------------------

    '//----------------------------------------------------------------------------------------------
    '//자격사항(Mat5)
    '//----------------------------------------------------------------------------------------------
    '//자격/면허
    Set oColumn = oMat5.Columns("License")

    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P125' AND U_UseYN= 'Y' ORDER BY CAST(U_Code AS NUMERIC) "
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        For i = 0 To oRecordSet.RecordCount - 1
            oColumn.ValidValues.Add oRecordSet.Fields(0).VALUE, oRecordSet.Fields(1).VALUE
            oRecordSet.MoveNext
        Next i
    End If
    oColumn.DisplayDesc = True

    '//----------------------------------------------------------------------------------------------
    '//발령사항(Mat6)
    '//----------------------------------------------------------------------------------------------
    '//발령구분
    Set oColumn = oMat6.Columns("appType")

    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P160' AND U_UseYN= 'Y' ORDER BY U_Code "
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        For i = 0 To oRecordSet.RecordCount - 1
            oColumn.ValidValues.Add oRecordSet.Fields(0).VALUE, oRecordSet.Fields(1).VALUE
            oRecordSet.MoveNext
        Next i
    End If
    oColumn.DisplayDesc = True

     '//사업장
    Set oColumn = oMat6.Columns("CLTCOD")
    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P144' AND U_UseYN= 'Y' ORDER BY CAST(U_Code AS NUMERIC) "
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        For i = 0 To oRecordSet.RecordCount - 1
            oColumn.ValidValues.Add oRecordSet.Fields(0).VALUE, oRecordSet.Fields(1).VALUE
            oRecordSet.MoveNext
        Next i
    End If
    oColumn.DisplayDesc = True

    '//부서
    Set oColumn = oMat6.Columns("TeamCode")
    sQry = " SELECT U_Code, U_CodeNm, U_Comment2 FROM [@PS_HR200L] WHERE Code='1' AND U_UseYN= 'Y' ORDER BY CAST(U_Code AS NUMERIC) "
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        For i = 0 To oRecordSet.RecordCount - 1
            oColumn.ValidValues.Add oRecordSet.Fields(0).VALUE, oRecordSet.Fields(1).VALUE & "(" & oRecordSet.Fields(2).VALUE & ")"
            oRecordSet.MoveNext
        Next i
    End If
    oColumn.DisplayDesc = True

    '//담당
    Set oColumn = oMat6.Columns("RspCode")
    sQry = " SELECT U_Code, U_CodeNm ,U_Comment2 FROM [@PS_HR200L] WHERE Code='2' AND U_UseYN= 'Y' ORDER BY CAST(U_Code AS NUMERIC) "
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        For i = 0 To oRecordSet.RecordCount - 1
            oColumn.ValidValues.Add oRecordSet.Fields(0).VALUE, oRecordSet.Fields(1).VALUE & "(" & oRecordSet.Fields(2).VALUE & ")"
            oRecordSet.MoveNext
        Next i
    End If
    oColumn.DisplayDesc = True


    '//직원구분
    Set oColumn = oMat6.Columns("JIGTYP")
    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P126' AND U_UseYN= 'Y' ORDER BY CAST(U_Code AS NUMERIC) "
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        For i = 0 To oRecordSet.RecordCount - 1
            oColumn.ValidValues.Add oRecordSet.Fields(0).VALUE, oRecordSet.Fields(1).VALUE
            oRecordSet.MoveNext
        Next i
    End If
    oColumn.DisplayDesc = True

    '//직책
    Set oColumn = oMat6.Columns("position")

    sQry = "SELECT posID,name FROM [OHPS] ORDER BY posID"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        For i = 0 To oRecordSet.RecordCount - 1
            oColumn.ValidValues.Add oRecordSet.Fields(0).VALUE, oRecordSet.Fields(1).VALUE
            oRecordSet.MoveNext
        Next i
    End If
    oColumn.DisplayDesc = True

    '// 호칭
    Set oColumn = oMat6.Columns("CallName")

    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P158' AND U_UseYN= 'Y' ORDER BY U_Code "
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        For i = 0 To oRecordSet.RecordCount - 1
            oColumn.ValidValues.Add oRecordSet.Fields(0).VALUE, oRecordSet.Fields(1).VALUE
            oRecordSet.MoveNext
        Next i
    End If
    oColumn.DisplayDesc = True

    '//직급
    Set oColumn = oMat6.Columns("JIGCOD")

    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P129' AND U_UseYN= 'Y' ORDER BY U_Code "
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        For i = 0 To oRecordSet.RecordCount - 1
            oColumn.ValidValues.Add oRecordSet.Fields(0).VALUE, oRecordSet.Fields(1).VALUE
            oRecordSet.MoveNext
        Next i
    End If
    oColumn.DisplayDesc = True

    '//----------------------------------------------------------------------------------------------
    '//동호회(Mat7)
    '//----------------------------------------------------------------------------------------------
    '//동호회
    Set oColumn = oMat7.Columns("Club")
    
    sQry = "           SELECT       U_Code,"
    sQry = sQry & "                 U_CodeNm"
    sQry = sQry & " FROM        [@PS_HR200L]"
    sQry = sQry & " WHERE       Code='P156'"
    sQry = sQry & "                 AND U_Char1 = '" & MDC_PS_Common.User_BPLID & "'"
    sQry = sQry & "                 AND U_UseYN= 'Y'"
    sQry = sQry & " ORDER BY   U_Seq"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        For i = 0 To oRecordSet.RecordCount - 1
            oColumn.ValidValues.Add oRecordSet.Fields(0).VALUE, oRecordSet.Fields(1).VALUE
            oRecordSet.MoveNext
        Next i
    End If
    oColumn.DisplayDesc = True

    '// 직책
'    Set oColumn = oMat7.Columns("Position")
'    sQry = "SELECT posID,name FROM [OHPS] ORDER BY posID"
'    oRecordSet.DoQuery sQry
'    If oRecordSet.RecordCount > 0 Then
'        For i = 0 To oRecordSet.RecordCount - 1
'            oColumn.ValidValues.Add oRecordSet.Fields(0).VALUE, oRecordSet.Fields(1).VALUE
'            oRecordSet.MoveNext
'        Next i
'    End If
'    oColumn.DisplayDesc = True
    '//----------------------------------------------------------------------------------------------
    '//학력사항(Mat8)
    '//----------------------------------------------------------------------------------------------
    '//학위
    Set oColumn = oMat8.Columns("degree")

    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P147' AND U_UseYN= 'Y' ORDER BY U_Code "
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        For i = 0 To oRecordSet.RecordCount - 1
            oColumn.ValidValues.Add oRecordSet.Fields(0).VALUE, oRecordSet.Fields(1).VALUE
            oRecordSet.MoveNext
        Next i
    End If
    oColumn.DisplayDesc = True

    '//이수구분
    Set oColumn = oMat8.Columns("complete")
    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P148' AND U_UseYN= 'Y' ORDER BY U_Code "
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        For i = 0 To oRecordSet.RecordCount - 1
            oColumn.ValidValues.Add oRecordSet.Fields(0).VALUE, oRecordSet.Fields(1).VALUE
            oRecordSet.MoveNext
        Next i
    End If
    oColumn.DisplayDesc = True
    '//----------------------------------------------------------------------------------------------
    '//교육사항(Mat9)
    '//----------------------------------------------------------------------------------------------
    '//분류
    Set oColumn = oMat9.Columns("type")
    sQry = "SELECT edType, Name FROM OHED"
    oRecordSet.DoQuery sQry

    If oRecordSet.RecordCount > 0 Then
        For i = 0 To oRecordSet.RecordCount - 1
            oColumn.ValidValues.Add oRecordSet.Fields(0).VALUE, oRecordSet.Fields(1).VALUE
            oRecordSet.MoveNext
        Next i
    End If
    oColumn.DisplayDesc = True


    '//----------------------------------------------------------------------------------------------
    '//상벌사항(Mat10)
    '//----------------------------------------------------------------------------------------------
    '//구분
    Set oColumn = oMat10.Columns("Category")
    oColumn.ValidValues.Add "1", "상"
    oColumn.ValidValues.Add "2", "벌"
    oColumn.DisplayDesc = True

    '//상벌종류
    Set oColumn = oMat10.Columns("Type")
    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P130' AND U_UseYN = 'Y' ORDER BY U_Code"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        For i = 0 To oRecordSet.RecordCount - 1
            oColumn.ValidValues.Add oRecordSet.Fields(0).VALUE, oRecordSet.Fields(1).VALUE
            oRecordSet.MoveNext
        Next i
    End If
    oColumn.DisplayDesc = True

    '//----------------------------------------------------------------------------------------------
    '//여권사항(Mat11)
    '//----------------------------------------------------------------------------------------------
    '구분
    Set oColumn = oMat11.Columns("TypeCode")
    oColumn.ValidValues.Add "1", "여권"
    oColumn.ValidValues.Add "2", "비자"
    
    '//여권종류
    Set oColumn = oMat11.Columns("PassType")
    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P151' AND U_UseYN= 'Y' ORDER BY U_Code "
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        For i = 0 To oRecordSet.RecordCount - 1
            oColumn.ValidValues.Add oRecordSet.Fields(0).VALUE, oRecordSet.Fields(1).VALUE
            oRecordSet.MoveNext
        Next i
    End If
    oColumn.DisplayDesc = True
    '//----------------------------------------------------------------------------------------------
    '//노조이력(Mat12)
    '//----------------------------------------------------------------------------------------------
    '//직책(노조)
    Set oColumn = oMat12.Columns("Position")
    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P152' AND U_UseYN= 'Y' ORDER BY U_Code "
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        For i = 0 To oRecordSet.RecordCount - 1
            oColumn.ValidValues.Add oRecordSet.Fields(0).VALUE, oRecordSet.Fields(1).VALUE
            oRecordSet.MoveNext
        Next i
    End If
    oColumn.DisplayDesc = True
    '//----------------------------------------------------------------------------------------------
    '//경조사항(Mat13)
    '//----------------------------------------------------------------------------------------------
    '//관계
    Set oColumn = oMat13.Columns("Relation")
    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P121' AND U_UseYN= 'Y' ORDER BY U_Code "
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        For i = 0 To oRecordSet.RecordCount - 1
            oColumn.ValidValues.Add oRecordSet.Fields(0).VALUE, oRecordSet.Fields(1).VALUE
            oRecordSet.MoveNext
        Next i
    End If
    oColumn.DisplayDesc = True
    '//----------------------------------------------------------------------------------------------
    '//사고현황(Mat14)
    '//----------------------------------------------------------------------------------------------
    '//사고구분
    Set oColumn = oMat14.Columns("type")
    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P153' AND U_UseYN= 'Y' ORDER BY U_Code "
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        For i = 0 To oRecordSet.RecordCount - 1
            oColumn.ValidValues.Add oRecordSet.Fields(0).VALUE, oRecordSet.Fields(1).VALUE
            oRecordSet.MoveNext
        Next i
    End If
    oColumn.DisplayDesc = True
    '//----------------------------------------------------------------------------------------------
    '//기타사항
    '//----------------------------------------------------------------------------------------------

    '// 주택유무
    Set oCombo = oForm.Items("House").Specific
    oCombo.ValidValues.Add "N", "No"
    oCombo.ValidValues.Add "Y", "Yes"
    oForm.Items("House").DisplayDesc = True

'    // 주거구분
    Set oCombo = oForm.Items("HType").Specific
    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P139' AND U_UseYN= 'Y' ORDER BY U_Code"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        Do Until oRecordSet.EOF
            oCombo.ValidValues.Add Trim$(oRecordSet.Fields(0).VALUE), Trim$(oRecordSet.Fields(1).VALUE)
            oRecordSet.MoveNext
        Loop
    End If
    oForm.Items("HType").DisplayDesc = True

'    //생활정도
    Set optBtn = oForm.Items("ClassJ").Specific
    optBtn.GroupWith ("ClassS")

    Set optBtn = oForm.Items("ClassH").Specific
    optBtn.GroupWith ("ClassS")

    '//군별
    Set oCombo = oForm.Items("military").Specific
    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P138' AND U_UseYN= 'Y' ORDER BY U_Code"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        Do Until oRecordSet.EOF
            oCombo.ValidValues.Add Trim$(oRecordSet.Fields(0).VALUE), Trim$(oRecordSet.Fields(1).VALUE)
            oRecordSet.MoveNext
        Loop
    End If
    oForm.Items("military").DisplayDesc = True

    '//병과
    Set oCombo = oForm.Items("Arm").Specific
    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P149' AND U_UseYN= 'Y' ORDER BY U_Code"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        Do Until oRecordSet.EOF
            oCombo.ValidValues.Add Trim$(oRecordSet.Fields(0).VALUE), Trim$(oRecordSet.Fields(1).VALUE)
            oRecordSet.MoveNext
        Loop
    End If
    oForm.Items("Arm").DisplayDesc = True

    '//계급
    Set oCombo = oForm.Items("Rank").Specific
    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P150' AND U_UseYN= 'Y' ORDER BY U_Code"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        Do Until oRecordSet.EOF
            oCombo.ValidValues.Add Trim$(oRecordSet.Fields(0).VALUE), Trim$(oRecordSet.Fields(1).VALUE)
            oRecordSet.MoveNext
        Loop
    End If
    oForm.Items("Rank").DisplayDesc = True

    '//역종
    Set oCombo = oForm.Items("StatusCd").Specific
    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P136' AND U_UseYN= 'Y' ORDER BY U_Code"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        Do Until oRecordSet.EOF
            oCombo.ValidValues.Add Trim$(oRecordSet.Fields(0).VALUE), Trim$(oRecordSet.Fields(1).VALUE)
            oRecordSet.MoveNext
        Loop
    End If
    oForm.Items("StatusCd").DisplayDesc = True

    '//전역구분
    Set oCombo = oForm.Items("DisType").Specific
    sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P137' AND U_UseYN= 'Y' ORDER BY U_Code"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        Do Until oRecordSet.EOF
            oCombo.ValidValues.Add Trim$(oRecordSet.Fields(0).VALUE), Trim$(oRecordSet.Fields(1).VALUE)
            oRecordSet.MoveNext
        Loop
    End If
    oForm.Items("DisType").DisplayDesc = True

    '//사내출입증
    Set oCombo = oForm.Items("military").Specific

    oCombo.ValidValues.Add "1", "전지역"
    oCombo.ValidValues.Add "2", "제한지역"
    oCombo.ValidValues.Add "3", "출입불가"

    oForm.Items("military").DisplayDesc = True

    '//종류 - 외국어 1
    Set oCombo = oForm.Items("foreLan1").Specific
    sQry = " SELECT Code, Name FROM [OLNG]  ORDER BY Code"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        Do Until oRecordSet.EOF
            oCombo.ValidValues.Add Trim$(oRecordSet.Fields(0).VALUE), Trim$(oRecordSet.Fields(1).VALUE)
            oRecordSet.MoveNext
        Loop
    End If
    oForm.Items("foreLan1").DisplayDesc = True

    '//종류 - 외국어 2
    Set oCombo = oForm.Items("foreLan2").Specific
    sQry = " SELECT Code, Name FROM [OLNG]  ORDER BY Code"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        Do Until oRecordSet.EOF
            oCombo.ValidValues.Add Trim$(oRecordSet.Fields(0).VALUE), Trim$(oRecordSet.Fields(1).VALUE)
            oRecordSet.MoveNext
        Loop
    End If
    oForm.Items("foreLan2").DisplayDesc = True

    '//혈액형1
    Set oCombo = oForm.Items("BloodTy1").Specific
    oCombo.ValidValues.Add "1", "RH+"
    oCombo.ValidValues.Add "2", "RH-"
    oForm.Items("BloodTy1").DisplayDesc = True


    '//혈액형2
    Set oCombo = oForm.Items("BloodTy2").Specific
    oCombo.ValidValues.Add "1", "A"
    oCombo.ValidValues.Add "2", "B"
    oCombo.ValidValues.Add "3", "O"
    oCombo.ValidValues.Add "4", "AB"
    oForm.Items("BloodTy2").DisplayDesc = True

    '//색맹여부
    Set oCombo = oForm.Items("colBlind").Specific
    oCombo.ValidValues.Add "N", "No"
    oCombo.ValidValues.Add "Y", "Yes"
    oForm.Items("colBlind").DisplayDesc = True

    '//검진구분
    Set oCombo = oForm.Items("MediExam").Specific
    oCombo.ValidValues.Add "1", "일반검진"
    oCombo.ValidValues.Add "9", "특수검진"

    oForm.Items("MediExam").DisplayDesc = True
    
    
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
    
PH_PY001_CreateItems_Error:
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY001_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY001_EnableMenus()
    
    On Error GoTo PH_PY001_EnableMenus_Error

    Call oForm.EnableMenu("1283", True)     '//제거
    Call oForm.EnableMenu("1284", False)    '//취소
    Call oForm.EnableMenu("1293", True)     '//행삭제
    
    Exit Sub
    
PH_PY001_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY001_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY001_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY001_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY001_FormItemEnabled
        Call PH_PY001_AddMatrixRow
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY001_FormItemEnabled
        oForm.Items("Code").Specific.VALUE = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY001_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY001_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY001_FormItemEnabled()
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim sQry        As String
    Dim i           As Long
    Dim optBtn      As SAPbouiCOM.OptionBtn
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    
    On Error GoTo PH_PY001_FormItemEnabled_Error
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Call oForm.Freeze(True)
    If (oForm.Mode = fm_ADD_MODE) Then
        'oForm.Items.Item("Code").Enabled = True
        oForm.Items.Item("FullName").Enabled = False
        oForm.Items.Item("CLTCOD").Enabled = True
        oForm.Items.Item("Btn01").Visible = False
        
        oDS_PH_PY001A.setValue "U_brthCntr", 0, "KR"    '//국적
        oDS_PH_PY001A.setValue "U_CldType", 0, "S"      '// 음양력
        
        '// 4.내국인구분
        Set oCombo = oForm.Items("INTGBN").Specific
        If oCombo.ValidValues.Count > 0 Then
            oCombo.Select 0, psk_Index
        End If
        '// 5.거주자구분
        Set oCombo = oForm.Items("DWEGBN").Specific
        If oCombo.ValidValues.Count > 0 Then
            oCombo.Select 0, psk_Index
        End If
        '// 6.외국인단일세율적용
        Set oCombo = oForm.Items("FRGTAX").Specific
        If oCombo.ValidValues.Count > 0 Then
            oCombo.Select 1, psk_Index
        End If
        
        oDS_PH_PY001A.setValue "U_GBHSEL", 0, "Y"
        oDS_PH_PY001A.setValue "U_BNSSEL", 0, "Y"
        oDS_PH_PY001A.setValue "U_TAXSEL", 0, "Y"
        oDS_PH_PY001A.setValue "U_JSNSEL", 0, "Y"
        oDS_PH_PY001A.setValue "U_Inform01", 0, "N"
        oDS_PH_PY001A.setValue "U_Inform02", 0, "N"
        oDS_PH_PY001A.setValue "U_Inform03", 0, "N"
        oDS_PH_PY001A.setValue "U_Inform04", 0, "N"
        oDS_PH_PY001A.setValue "U_Inform05", 0, "N"
        oDS_PH_PY001A.setValue "U_Inform06", 0, "N"
        
        
        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", False)    '//문서추가

        '// 접속자에 따른 사업장 선택
'        Call oDS_PH_PY001A.setValue("U_CLTCOD", 0, MDC_SetMod.Get_ReData("Branch", "USER_CODE", "OUSR", "'" & oCompany.UserName & "'"))
        
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        
        '//기본사항 - 부서 (사업장에 따른 부서변경)
        Set oCombo = oForm.Items("TeamCode").Specific

        If oCombo.ValidValues.Count > 0 Then
            For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                oCombo.ValidValues.Remove i, psk_Index
            Next i
            oCombo.ValidValues.Add "", ""
            oCombo.Select 0, psk_Index
        End If
        
        If oDS_PH_PY001A.GetValue("U_CLTCOD", 0) <> "" Then
            sQry = "           SELECT       '' AS [Code], "
            sQry = sQry & "                 '' AS [Name],"
            sQry = sQry & "                 -1 AS [Seq]"
            sQry = sQry & " UNION ALL"
            sQry = sQry & " SELECT       U_Code AS [Code], "
            sQry = sQry & "                 U_CodeNm AS [Name],"
            sQry = sQry & "                 U_Seq AS [Seq]"
            sQry = sQry & " FROM        [@PS_HR200L] "
            sQry = sQry & " WHERE       Code = '1' "
            sQry = sQry & "                 AND U_Char2 = '" & oForm.Items.Item("CLTCOD").Specific.VALUE & "'"
            sQry = sQry & "                 AND U_UseYN = 'Y'"
            sQry = sQry & " ORDER BY   Seq"
        
        
'            sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
'            sQry = sQry & " WHERE Code = '1' AND U_Char2 = '" & oForm.Items.Item("CLTCOD").Specific.Value & "'"
'            sQry = sQry & " ORDER BY U_Code"
            Call SetReDataCombo(oForm, sQry, oCombo)
        End If
        
        '//담당 (사업장에 따른 담당변경)
    
        Set oCombo = oForm.Items("RspCode").Specific

        If oCombo.ValidValues.Count > 0 Then
            For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                oCombo.ValidValues.Remove i, psk_Index
            Next i
            oCombo.ValidValues.Add "", ""
            oCombo.Select 0, psk_Index
        End If
        
        If oForm.Items.Item("CLTCOD").Specific.VALUE <> "" Then
            sQry = "           SELECT       '' AS [Code], "
            sQry = sQry & "                 '' AS [Name],"
            sQry = sQry & "                 -1 AS [Seq]"
            sQry = sQry & " UNION ALL"
            sQry = sQry & " SELECT       U_Code AS [Code], "
            sQry = sQry & "                 U_CodeNm AS [Name],"
            sQry = sQry & "                 U_Seq AS [Seq]"
            sQry = sQry & " FROM        [@PS_HR200L] "
            sQry = sQry & " WHERE       Code = '2' "
            sQry = sQry & "                 AND U_Char2 = '" & oForm.Items.Item("CLTCOD").Specific.VALUE & "'"
            sQry = sQry & "                 AND U_Char1 = '" & oForm.Items.Item("TeamCode").Specific.VALUE & "'"
            sQry = sQry & "                 AND U_UseYN = 'Y'"
            sQry = sQry & " ORDER BY   Seq"
        
'            sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
'            sQry = sQry & " WHERE Code = '2' AND U_Char2 = '" & oForm.Items.Item("CLTCOD").Specific.Value & "'"
'            sQry = sQry & " Order By U_Code"
            Call SetReDataCombo(oForm, sQry, oCombo)
        End If
        
        '//반 (사업장에 따른 반변경)
        
        Set oCombo = oForm.Items("ClsCode").Specific
        
        If oCombo.ValidValues.Count > 0 Then
            For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                oCombo.ValidValues.Remove i, psk_Index
            Next i
            oCombo.ValidValues.Add "", ""
            oCombo.Select 0, psk_Index
        End If
        
        If oForm.Items.Item("CLTCOD").Specific.VALUE <> "" Then
            sQry = "           SELECT       '' AS [Code], "
            sQry = sQry & "                 '' AS [Name],"
            sQry = sQry & "                 -1 AS [Seq]"
            sQry = sQry & " UNION ALL"
            sQry = sQry & " SELECT       U_Code AS [Code], "
            sQry = sQry & "                 U_CodeNm AS [Name],"
            sQry = sQry & "                 U_Seq AS [Seq]"
            sQry = sQry & " FROM        [@PS_HR200L] "
            sQry = sQry & " WHERE       Code = '9' "
            sQry = sQry & "                 AND U_Char3 = '" & oForm.Items.Item("CLTCOD").Specific.VALUE & "'"
            sQry = sQry & "                 AND U_Char2 = '" & oForm.Items.Item("TeamCode").Specific.VALUE & "'"
            sQry = sQry & "                 AND U_Char1 = '" & oForm.Items.Item("RspCode").Specific.VALUE & "'"
            sQry = sQry & "                 AND U_UseYN = 'Y'"
            sQry = sQry & " ORDER BY   Seq"
        
'            sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
'            sQry = sQry & " WHERE Code='9' AND U_Char3 = '" & oForm.Items.Item("CLTCOD").Specific.Value & "'"
'            sQry = sQry & " ORDER BY U_Code"
            Call SetReDataCombo(oForm, sQry, oCombo)
        End If
        '//커서를 첫번째 ITEM으로 지정
        oForm.ActiveItem = "Code"
    ElseIf (oForm.Mode = fm_FIND_MODE) Then
    
        '사진 이미지 초기화(2017.05.25 송명규)
        oForm.Items("Pic1").Specific.Picture = ""
    
        'oForm.Items.Item("Code").Enabled = True
        oForm.Items.Item("FullName").Enabled = True
        oForm.Items.Item("CLTCOD").Enabled = True
        oForm.Items.Item("Btn01").Visible = True
        
        '// 4.내국인구분
        Set oCombo = oForm.Items("INTGBN").Specific
        If oCombo.ValidValues.Count > 0 Then
            oCombo.Select 0, psk_Index
        End If
        '// 5.거주자구분
        Set oCombo = oForm.Items("DWEGBN").Specific
        If oCombo.ValidValues.Count > 0 Then
            oCombo.Select 0, psk_Index
        End If
        '// 6.외국인단일세율적용
        Set oCombo = oForm.Items("FRGTAX").Specific
        If oCombo.ValidValues.Count > 0 Then
            oCombo.Select 1, psk_Index
        End If
        
        oDS_PH_PY001A.setValue "U_GBHSEL", 0, "Y"
        oDS_PH_PY001A.setValue "U_BNSSEL", 0, "Y"
        oDS_PH_PY001A.setValue "U_TAXSEL", 0, "Y"
        oDS_PH_PY001A.setValue "U_JSNSEL", 0, "Y"
        
        Call oForm.EnableMenu("1281", False)    '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가
        
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        
'        '//부서
'        Set oCombo = oForm.Items("TeamCode").Specific
'        If oCombo.ValidValues.Count > 0 Then
'            For i = oCombo.ValidValues.Count - 1 To 0 Step -1
'                oCombo.ValidValues.Remove i, psk_Index
'            Next i
'            oCombo.ValidValues.Add "", "-"
'            oCombo.Select 0, psk_Index
'        End If
'
'        If oDS_PH_PY001A.GetValue("U_CLTCOD", 0) <> "" Then
'            sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
'            sQry = sQry & " WHERE Code = '1' AND U_Char2 = '" & oForm.Items.Item("CLTCOD").Specific.Value & "'"
'            sQry = sQry & " ORDER BY U_Code"
'            Call SetReDataCombo(oForm, sQry, oCombo)
'        End If
'
'        '//담당
'        Set oCombo = oForm.Items("RspCode").Specific
'        If oCombo.ValidValues.Count > 0 Then
'            For i = oCombo.ValidValues.Count - 1 To 0 Step -1
'                oCombo.ValidValues.Remove i, psk_Index
'            Next i
'            oCombo.ValidValues.Add "", "-"
'            oCombo.Select 0, psk_Index
'        End If
'
'        If oForm.Items.Item("CLTCOD").Specific.Value <> "" Then
'            sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
'            sQry = sQry & " WHERE Code = '2' AND U_Char2 = '" & oForm.Items.Item("CLTCOD").Specific.Value & "'"
'            sQry = sQry & " Order By U_Code"
'            Call SetReDataCombo(oForm, sQry, oCombo)
'        End If
'
'
'
'        '//반 (사업장에 따른 반변경)
'
'        Set oCombo = oForm.Items("ClsCode").Specific
'
'        If oCombo.ValidValues.Count > 0 Then
'            For i = oCombo.ValidValues.Count - 1 To 0 Step -1
'                oCombo.ValidValues.Remove i, psk_Index
'            Next i
'            oCombo.ValidValues.Add "", "-"
'            oCombo.Select 0, psk_Index
'        End If
        
        '//근무조
        Set oCombo = oForm.Items("GNMUJO").Specific
        If oCombo.ValidValues.Count > 0 Then
            For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                oCombo.ValidValues.Remove i, psk_Index
            Next i
            oCombo.ValidValues.Add "", "-"
            oCombo.Select 0, psk_Index
        End If
        
    ElseIf (oForm.Mode = fm_OK_MODE) Then
        'oForm.ActiveItem = "1"
        'oForm.Items.Item("Code").Enabled = False
        oForm.Items.Item("FullName").Enabled = False
        oForm.Items.Item("CLTCOD").Enabled = False
        
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD", False)
        
        If oForm.Mode = fm_FIND_MODE Then
            Call oForm.EnableMenu("1281", False)     '//문서찾기
        Else
            Call oForm.EnableMenu("1281", True)     '//문서찾기
        End If
        
        'Call oForm.EnableMenu("1281", False)     '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가
    
        
    
    End If
    Set optBtn = Nothing
    Set oCombo = Nothing
    Call oForm.Freeze(False)
    Exit Sub
    
PH_PY001_FormItemEnabled_Error:
    Set optBtn = Nothing
    Set oCombo = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY001_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Public Sub Raise_FormItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
    Dim sQry        As String
    Dim i           As Long
    Dim tSex        As String
    Dim tBrith      As String
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns     As SAPbouiCOM.Columns
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo Raise_FormItemEvent_Error
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Select Case pval.EventType
        Case et_ITEM_PRESSED: '//
            If pval.BeforeAction = True Then
                If pval.ItemUID = "1" Then
                    If oForm.Mode = fm_ADD_MODE Or oForm.Mode = fm_UPDATE_MODE Then
                        If PH_PY001_DataValidCheck = False Then
                            BubbleEvent = False
                        End If
                    End If
                End If
                
                If oForm.Mode = fm_FIND_MODE Then
                    If pval.ItemUID = "Btn01" Then
                        oForm.Items("Code").CLICK ct_Regular
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                End If
                If pval.ItemUID = "Btn02" Then
                    oForm.Items("HOBONG").CLICK ct_Regular
                    Sbo_Application.ActivateMenuItem ("7425")
                    BubbleEvent = False
                End If
                If pval.ItemUID = "Pic2" Then
                    If oForm.Mode <> fm_ADD_MODE Or oForm.Mode <> fm_FIND_MODE Then
                        Call PH_PY001_LoadPic
                    End If
                End If
                If pval.ItemUID = "Pic3" Then
                    oDS_PH_PY001A.setValue "U_Pic1", 0, ""
                    If oForm.Mode = fm_OK_MODE Then
                        oForm.Mode = fm_UPDATE_MODE
                    End If
                End If
            ElseIf pval.BeforeAction = False Then
                Select Case pval.ItemUID
                Case "FLD01", "FLD02", "FLD03", "FLD04", "FLD05", "FLD06", "FLD07", "FLD08", "FLD09", "FLD10", "FLD11", "FLD12", "FLD13", "FLD14", "FLD15", "FLD16"
                    oForm.PaneLevel = Mid(pval.ItemUID, 4)
                
                Case "1"
                    If oForm.Mode = fm_ADD_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY001_empIDCheck
                            Call PH_PY001_FormItemEnabled
                            Call PH_PY001_AddMatrixRow
                        End If
                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY001_FormItemEnabled
                            Call PH_PY001_AddMatrixRow
                        End If
                    ElseIf oForm.Mode = fm_OK_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY001_FormItemEnabled
                        End If
                    ElseIf oForm.Mode = fm_FIND_MODE Then
                        If pval.ActionSuccess = True Then
                            oForm.Mode = fm_OK_MODE
                            Call PH_PY001_FormItemEnabled
                        End If
                    End If
                End Select
                
            End If
     
'----------------------------------------------------------
        Case et_KEY_DOWN: '//2

'----------------------------------------------------------
        Case et_GOT_FOCUS: '//3
            Select Case pval.ItemUID
            Case "Mat1", "Mat2", "Mat3", "Mat4", "Mat5", "Mat6", "Mat7", "Mat8", "Mat9", "Mat10", "Mat11", "Mat12", "Mat13", "Mat14", "Mat15"
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
                    If pval.ItemUID = "CLTCOD" Then
                        
                        '//기본사항 - 부서 (사업장에 따른 부서변경)
                        Set oCombo = oForm.Items("TeamCode").Specific
    
                        If oCombo.ValidValues.Count > 0 Then
                            For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                                oCombo.ValidValues.Remove i, psk_Index
                            Next i
                            oCombo.ValidValues.Add "", ""
                            oCombo.Select 0, psk_Index
                        End If
                        
                        sQry = "           SELECT       '' AS [Code], "
                        sQry = sQry & "                 '' AS [Name],"
                        sQry = sQry & "                 -1 AS [Seq]"
                        sQry = sQry & " UNION ALL"
                        sQry = sQry & " SELECT       U_Code AS [Code], "
                        sQry = sQry & "                 U_CodeNm AS [Name],"
                        sQry = sQry & "                 U_Seq AS [Seq]"
                        sQry = sQry & " FROM        [@PS_HR200L] "
                        sQry = sQry & " WHERE       Code = '1' "
                        sQry = sQry & "                 AND U_Char2 = '" & oForm.Items.Item("CLTCOD").Specific.VALUE & "'"
                        sQry = sQry & "                 AND U_UseYN = 'Y'"
                        sQry = sQry & " ORDER BY   Seq"
                        Call SetReDataCombo(oForm, sQry, oCombo)
                        
                        oForm.Items("TeamCode").DisplayDesc = True
                        
                        '//담당 (팀에 따른 담당변경)
                    ElseIf pval.ItemUID = "TeamCode" Then
                    
                        Set oCombo = oForm.Items("RspCode").Specific
    
                        If oCombo.ValidValues.Count > 0 Then
                            For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                                oCombo.ValidValues.Remove i, psk_Index
                            Next i
                            oCombo.ValidValues.Add "", ""
                            oCombo.Select 0, psk_Index
                        End If
                        
                        sQry = "           SELECT       '' AS [Code], "
                        sQry = sQry & "                 '' AS [Name],"
                        sQry = sQry & "                 -1 AS [Seq]"
                        sQry = sQry & " UNION ALL"
                        sQry = sQry & " SELECT       U_Code AS [Code], "
                        sQry = sQry & "                 U_CodeNm AS [Name],"
                        sQry = sQry & "                 U_Seq AS [Seq]"
                        sQry = sQry & " FROM        [@PS_HR200L] "
                        sQry = sQry & " WHERE       Code = '2' "
                        sQry = sQry & "                 AND U_Char2 = '" & oForm.Items.Item("CLTCOD").Specific.VALUE & "'"
                        sQry = sQry & "                 AND U_Char1 = '" & oForm.Items.Item("TeamCode").Specific.VALUE & "'"
                        sQry = sQry & "                 AND U_UseYN = 'Y'"
                        sQry = sQry & " ORDER BY   Seq"
                        
                        
'                        sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
'                        sQry = sQry & " WHERE Code = '2' AND U_Char2 = '" & oForm.Items.Item("CLTCOD").Specific.Value & "'"
'                        sQry = sQry & " Order By U_Code"
                        Call SetReDataCombo(oForm, sQry, oCombo)
                        
                        oForm.Items("RspCode").DisplayDesc = True
                        
                        '//반 (사업장에 따른 반변경)
                    ElseIf pval.ItemUID = "RspCode" Then
                    
                        Set oCombo = oForm.Items("ClsCode").Specific
                        
                        If oCombo.ValidValues.Count > 0 Then
                            For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                                oCombo.ValidValues.Remove i, psk_Index
                            Next i
                            oCombo.ValidValues.Add "", ""
                            oCombo.Select 0, psk_Index
                        End If
                        
                        sQry = "           SELECT       '' AS [Code], "
                        sQry = sQry & "                 '' AS [Name],"
                        sQry = sQry & "                 -1 AS [Seq]"
                        sQry = sQry & " UNION ALL"
                        sQry = sQry & " SELECT       U_Code AS [Code], "
                        sQry = sQry & "                 U_CodeNm AS [Name],"
                        sQry = sQry & "                 U_Seq AS [Name]"
                        sQry = sQry & " FROM        [@PS_HR200L] "
                        sQry = sQry & " WHERE       Code = '9' "
                        sQry = sQry & "                 AND U_Char3 = '" & oForm.Items.Item("CLTCOD").Specific.VALUE & "'"
                        sQry = sQry & "                 AND U_Char2 = '" & oForm.Items.Item("TeamCode").Specific.VALUE & "'"
                        sQry = sQry & "                 AND U_Char1 = '" & oForm.Items.Item("RspCode").Specific.VALUE & "'"
                        sQry = sQry & "                 AND U_UseYN = 'Y'"
                        sQry = sQry & " ORDER BY   Seq"
                        
'                        sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
'                        sQry = sQry & " WHERE Code='9' AND U_Char3 = '" & oForm.Items.Item("CLTCOD").Specific.Value & "'"
'                        sQry = sQry & " ORDER BY U_Code"
                        
                        Call SetReDataCombo(oForm, sQry, oCombo)
                        
                        oForm.Items("GrpDat").CLICK ct_Regular
                    End If
                    
                    '//재직구분에 따른 퇴직일자 활성화
                    If pval.ItemUID = "status" Then
                        Select Case oDS_PH_PY001A.GetValue("U_status", 0)
                        Case 1, 2, 3, 4
                            Call oDS_PH_PY001A.setValue("U_termDate", 0, "")
                            oForm.Items("termDate").Enabled = False
                        Case 5
                            oForm.Items("termDate").Enabled = True
                        End Select
                    End If
                    
                    '//혼인구분에 따른 결혼기념일 활성화
                    If pval.ItemUID = "martStat" Then
                        Select Case oDS_PH_PY001A.GetValue("U_martStat", 0)
                        Case "N"
                            Call oDS_PH_PY001A.setValue("U_weddingD", 0, "")
                            oForm.Items("weddingD").Enabled = False
                            
                        Case "Y"
                            oForm.Items("weddingD").Enabled = True
                        End Select
                    End If
                    
                    '//근무형태에 따른 근무조 값
                    If pval.ItemUID = "ShiftDat" Then
                        Set oCombo = oForm.Items("GNMUJO").Specific
    
                        If oCombo.ValidValues.Count > 0 Then
                            For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                                oCombo.ValidValues.Remove i, psk_Index
                            Next i
                        End If
                        
                        sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
                        sQry = sQry & " WHERE Code = 'P155' AND U_Char1 = '" & oForm.Items.Item("ShiftDat").Specific.VALUE & "'"
                        sQry = sQry & " ORDER BY U_Code"
                        Call SetReDataCombo(oForm, sQry, oCombo)
                        
                        oCombo.Select 0, psk_Index
                        oForm.Items("GNMUJO").DisplayDesc = True
                    End If
                    If pval.ItemUID = "PAYTYP" Then
                        Select Case Trim(oDS_PH_PY001A.GetValue("U_PAYTYP", 0))
                        Case "1"
                            Call oDS_PH_PY001A.setValue("U_HOBONG", 0, "000")
                            oForm.Items("Btn02").Visible = False
                            oForm.Items("HOBONG").Enabled = False
                            oForm.Items("YUNAMT").Enabled = True
                        Case "2"
                            Call oDS_PH_PY001A.setValue("U_HOBONG", 0, "")
                            oForm.Items("Btn02").Visible = True
                            oForm.Items("HOBONG").Enabled = True
                            oForm.Items("YUNAMT").Enabled = False
                        Case "3"
                            Call oDS_PH_PY001A.setValue("U_HOBONG", 0, "000")
                            oForm.Items("Btn02").Visible = False
                            oForm.Items("HOBONG").Enabled = False
                            oForm.Items("YUNAMT").Enabled = False
                        Case "4"
                            Call oDS_PH_PY001A.setValue("U_HOBONG", 0, "000")
                            oForm.Items("Btn02").Visible = False
                            oForm.Items("HOBONG").Enabled = False
                            oForm.Items("YUNAMT").Enabled = False
                        End Select
                        oForm.Items("HOBONG").Update
                    End If
                End If
            End If

            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_CLICK: '//6
            If pval.BeforeAction = True Then
                Select Case pval.ItemUID
                Case "Mat1", "Mat2", "Mat3", "Mat4", "Mat5", "Mat6", "Mat7", "Mat8", "Mat9", "Mat10", "Mat11", "Mat12", "Mat13", "Mat14", "Mat15"
                    If pval.Row > 0 Then
                        Select Case pval.ItemUID
                        Case "Mat1"
                            Call oMat1.SelectRow(pval.Row, True, False)
                        Case "Mat2"
                            Call oMat2.SelectRow(pval.Row, True, False)
                        Case "Mat3"
                            Call oMat3.SelectRow(pval.Row, True, False)
                        Case "Mat4"
                            Call oMat4.SelectRow(pval.Row, True, False)
                        Case "Mat5"
                            Call oMat5.SelectRow(pval.Row, True, False)
                        Case "Mat6"
                            Call oMat6.SelectRow(pval.Row, True, False)
                        Case "Mat7"
                            Call oMat7.SelectRow(pval.Row, True, False)
                        Case "Mat8"
                            Call oMat8.SelectRow(pval.Row, True, False)
                        Case "Mat9"
                            Call oMat9.SelectRow(pval.Row, True, False)
                        Case "Mat10"
                            Call oMat10.SelectRow(pval.Row, True, False)
                        Case "Mat11"
                            Call oMat11.SelectRow(pval.Row, True, False)
                        Case "Mat12"
                            Call oMat12.SelectRow(pval.Row, True, False)
                        Case "Mat13"
                            Call oMat13.SelectRow(pval.Row, True, False)
                        Case "Mat14"
                            Call oMat14.SelectRow(pval.Row, True, False)
                        Case "Mat15"
                            Call oMat15.SelectRow(pval.Row, True, False)
                    
                        End Select
                    End If
                End Select
                
                Select Case pval.ItemUID
                Case "Mat1", "Mat2", "Mat3", "Mat4", "Mat5", "Mat6", "Mat7", "Mat8", "Mat9", "Mat10", "Mat11", "Mat12", "Mat13", "Mat14", "Mat15"
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
                If pval.ItemUID = "Mat3" And pval.ColUID = "FamPer" Then
                    If oMat3.Columns("FamPer").Cells(pval.Row).Specific.VALUE <> "" Then
                        If GovIDCheck(oMat3.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE) = False Then
                            Sbo_Application.SetStatusBarMessage "가족사항 - 잘못된 주민번호입니다.", bmt_Short, True
                            BubbleEvent = False
'                        Else
'                            If Mid(oMat3.Columns("FamPer").Cells(pval.Row).Specific.Value, 7, 1) = 1 Or Mid(oMat3.Columns("FamPer").Cells(pval.Row).Specific.Value, 7, 1) = 2 Then
'                                oMat3.Columns("BIRDAT").Cells(pval.Row).Specific.Value = Format(Val(Left(oMat3.Columns("FamPer").Cells(pval.Row).Specific.Value, 6)), "19######")
'                            ElseIf Mid(oMat3.Columns("FamPer").Cells(pval.Row).Specific.Valu, 7, 1) = 3 Or Mid(oMat3.Columns("FamPer").Cells(pval.Row).Specific.Valu, 7, 1) = 4 Then
'                                oMat3.Columns("BIRDAT").Cells(pval.Row).Specific.Value = Format(Val(Left(oMat3.Columns("FamPer").Cells(pval.Row).Specific.Value, 6)), "20######")
'                            End If
                        End If
                    End If
                End If
            
                If pval.ItemUID = "govID" Then
                    If oDS_PH_PY001A.GetValue("U_govID", 0) <> "" Then
                        If GovIDCheck(oDS_PH_PY001A.GetValue("U_govID", 0)) = True Then
                            '// 성별구분
                            If Mid(oDS_PH_PY001A.GetValue("U_govID", 0), 7, 1) = 1 Or Mid(oDS_PH_PY001A.GetValue("U_govID", 0), 7, 1) = 3 Then '
                                tSex = "M"
                            ElseIf Mid(oDS_PH_PY001A.GetValue("U_govID", 0), 7, 1) = 2 Or Mid(oDS_PH_PY001A.GetValue("U_govID", 0), 7, 1) = 4 Then
                                tSex = "F"
                            End If
                            
                            Call oDS_PH_PY001A.setValue("U_Sex", 0, tSex)
                            
'                            '// 생년월일
'                            If Mid(oDS_PH_PY001A.GetValue("U_govID", 0), 7, 1) = 1 Or Mid(oDS_PH_PY001A.GetValue("U_govID", 0), 7, 1) = 2 Then
'                                Call oDS_PH_PY001A.setValue("U_birthDat", 0, Format(Mid(oDS_PH_PY001A.GetValue("U_govID", 0), 1, 6), "19######"))
'                            ElseIf Mid(oDS_PH_PY001A.GetValue("U_govID", 0), 7, 1) = 3 Or Mid(oDS_PH_PY001A.GetValue("U_govID", 0), 7, 1) = 4 Then
'                                Call oDS_PH_PY001A.setValue("U_birthDat", 0, Format(Mid(oDS_PH_PY001A.GetValue("U_govID", 0), 1, 6), "20######"))
'                            End If
                        Else
                            Sbo_Application.SetStatusBarMessage "잘못된 주민번호입니다" & Err.Description, bmt_Short, True
                            BubbleEvent = False
                        End If
                    End If
                End If
'                If pval.ItemUID = "birthDat" Then
'                    If oDS_PH_PY001A.GetValue("U_birthDat", 0) = "" Then
'                        Sbo_Application.SetStatusBarMessage "생년월일은 필수 입니다" & Err.Description, bmt_Short, True
'                        BubbleEvent = False
'                    End If
'                End If
                    
            ElseIf pval.BeforeAction = False Then
                If pval.ItemChanged = True Then
                    Select Case pval.ItemUID
'                    Case "Code"
'
'                    Case "Name"
'                        Call PH_PY001_FlushToItemValue(pval.ItemUID)
                    Case "FullName"
                        If oForm.Items("FullName").Specific.VALUE <> "" Then
                            Call oDS_PH_PY001A.setValue("U_MZBURI", 0, IIf(Trim(oDS_PH_PY001A.GetValue("U_sex", 0)) = "F", "1", "0"))
                        End If
                    Case "HOBONG"      '//호봉코드
                        If oForm.Items("HOBONG").Specific.VALUE = "" Then
                            oDS_PH_PY001A.setValue "U_YUNAMT", 0, 0
                            oDS_PH_PY001A.setValue "U_STDAMT", 0, 0
                            oDS_PH_PY001A.setValue "U_BNSAMT", 0, 0
                            Call PH_PY001_Matrix_Inti
                        Else
                            Call PH_PY001_Display_Hobong
                        End If
                    Case "YUNAMT"   '//연봉
                        If Val(Trim(oDS_PH_PY001A.GetValue("U_YUNAMT", 0))) > 0 And Trim(oDS_PH_PY001A.GetValue("U_PAYTYP", 0)) = "1" Then
                            oDS_PH_PY001A.setValue "U_STDAMT", 0, Val(Trim(oDS_PH_PY001A.GetValue("U_YUNAMT", 0))) / 20
                            oDS_PH_PY001A.setValue "U_BNSAMT", 0, Val(Trim(oDS_PH_PY001A.GetValue("U_YUNAMT", 0))) / 20
                        End If
                        
                    Case "KUKGRD"   '//건강보험등급
                        If Len(oForm.Items(pval.ItemUID).Specific.String) > 0 Then
                            oDS_PH_PY001A.setValue "U_KUKGRD", 0, Format$(oForm.Items(pval.ItemUID).Specific.String, "000")
                        End If
                    Case "BUYN20", "BUYN60"   '//부양가족20세이하 , 부양가족60세이상
                        oDS_PH_PY001A.setValue "U_BUYNSU", 0, Val(oDS_PH_PY001A.GetValue("U_BUYN20", 0)) + Val(oDS_PH_PY001A.GetValue("U_BUYN60", 0))
                        
                    Case "DAGYSU"   '//다자녀
                        If Val(oDS_PH_PY001A.GetValue("U_BUYN20", 0)) < Val(oDS_PH_PY001A.GetValue("U_DAGYSU", 0)) Then
                            oDS_PH_PY001A.setValue "U_BUYN20", 0, Val(oDS_PH_PY001A.GetValue("U_DAGYSU", 0))
                            '//부양가족수
                            oDS_PH_PY001A.setValue "U_BUYNSU", 0, Val(oDS_PH_PY001A.GetValue("U_BUYN20", 0)) + Val(oDS_PH_PY001A.GetValue("U_BUYN60", 0))
                        End If
                    Case "GBHAMT"   '//고용보험보수월액
                        If Val(oDS_PH_PY001A.GetValue("U_GBHAMT", 0)) = 0 Then
                            oDS_PH_PY001A.setValue "U_GBHSEL", 0, "N"
                        Else
                            oDS_PH_PY001A.setValue "U_GBHSEL", 0, "Y"
                        End If
                        
                    '/매트릭스/---------------------------------------------------------------------------------------
                    Case "Mat1"
                        If pval.ColUID = "FILD01" Then
'                            oDS_PH_PY001B.setValue "U_FILD02", pval.Row, MDC_SetMod.Get_ReData("CardName", "CardCode", "OCRD", "'" & oDS_PH_PY001B.GetValue("COL1", pval.Row) & "'", "")
                            Call PH_PY001_AddMatrixRow
                            Call oMat1.Columns(pval.ColUID).Cells(pval.Row).CLICK(ct_Regular)
                        End If
                        
                    Case "Mat2"
                        If pval.ColUID = "FILD01" Then
'                            oDS_PH_PY001C.setValue "U_FILD02", pval.Row, MDC_SetMod.Get_ReData("CardName", "CardCode", "OCRD", "'" & oDS_PH_PY001C.GetValue("COL1", pval.Row) & "'", "")
                            Call PH_PY001_AddMatrixRow
                            Call oMat2.Columns(pval.ColUID).Cells(pval.Row).CLICK(ct_Regular)
                        End If
                        
                    Case "Mat3"
                        If pval.ColUID = "FamNam" Then
                            Call PH_PY001_AddMatrixRow
                            Call oMat3.Columns(pval.ColUID).Cells(pval.Row).CLICK(ct_Regular)
                        End If
                        
                    Case "Mat4"
                        If pval.ColUID = "employer" Then
                            Call PH_PY001_AddMatrixRow
                            Call oMat4.Columns(pval.ColUID).Cells(pval.Row).CLICK(ct_Regular)
                        End If
                        
                    Case "Mat5"
                        If pval.ColUID = "LCNumber" Then
                            Call PH_PY001_AddMatrixRow
                            Call oMat5.Columns(pval.ColUID).Cells(pval.Row).CLICK(ct_Regular)
                        End If
                        
                    Case "Mat6"
                        If pval.ColUID = "appNum" Then
                            Call PH_PY001_AddMatrixRow
                            Call oMat6.Columns(pval.ColUID).Cells(pval.Row).CLICK(ct_Regular)
                        End If
                        
                    Case "Mat7"
                        If pval.ColUID = "InDate" Then
                            Call PH_PY001_AddMatrixRow
                            Call oMat7.Columns(pval.ColUID).Cells(pval.Row).CLICK(ct_Regular)
                        End If
                        
                    Case "Mat8"
                        If pval.ColUID = "School" Then
                            Call PH_PY001_AddMatrixRow
                            Call oMat8.Columns(pval.ColUID).Cells(pval.Row).CLICK(ct_Regular)
                        End If
                        
                    Case "Mat9"
                        If pval.ColUID = "major" Then
                            Call PH_PY001_AddMatrixRow
                            Call oMat9.Columns(pval.ColUID).Cells(pval.Row).CLICK(ct_Regular)
                        End If
                        
                    Case "Mat10"
                        If pval.ColUID = "Basis" Then
                            Call PH_PY001_AddMatrixRow
                            Call oMat10.Columns(pval.ColUID).Cells(pval.Row).CLICK(ct_Regular)
                        End If
                        
                    Case "Mat11"
                        If pval.ColUID = "Passport" Then
                            Call PH_PY001_AddMatrixRow
                            Call oMat11.Columns(pval.ColUID).Cells(pval.Row).CLICK(ct_Regular)
                        End If
                        
                    Case "Mat12"
                        If pval.ColUID = "Period" Then
                            Call PH_PY001_AddMatrixRow
                            Call oMat12.Columns(pval.ColUID).Cells(pval.Row).CLICK(ct_Regular)
                        End If
                        
                    Case "Mat13"
                        If pval.ColUID = "Person" Then
                            Call PH_PY001_AddMatrixRow
                            Call oMat13.Columns(pval.ColUID).Cells(pval.Row).CLICK(ct_Regular)
                        End If
                        
                    Case "Mat14"
                        If pval.ColUID = "injury" Then
                            Call PH_PY001_AddMatrixRow
                            Call oMat14.Columns(pval.ColUID).Cells(pval.Row).CLICK(ct_Regular)
                        End If
                    Case "Mat15"
                        If pval.ColUID = "Car" Then
                            Call PH_PY001_AddMatrixRow
                            Call oMat15.Columns(pval.ColUID).Cells(pval.Row).CLICK(ct_Regular)
                        End If
                        
                    End Select
                 
                End If
            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_MATRIX_LOAD: '//11
            If pval.BeforeAction = True Then
            ElseIf pval.BeforeAction = False Then
                oMat1.LoadFromDataSource
                oMat2.LoadFromDataSource
                oMat3.LoadFromDataSource
                oMat4.LoadFromDataSource
                oMat5.LoadFromDataSource
                oMat6.LoadFromDataSource
                oMat7.LoadFromDataSource
                oMat8.LoadFromDataSource
                oMat9.LoadFromDataSource
                oMat10.LoadFromDataSource
                oMat11.LoadFromDataSource
                oMat12.LoadFromDataSource
                oMat13.LoadFromDataSource
                oMat14.LoadFromDataSource
                oMat15.LoadFromDataSource
                Call PH_PY001_AddMatrixRow
                
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
                Set oDS_PH_PY001A = Nothing
                Set oDS_PH_PY001B = Nothing
                Set oDS_PH_PY001C = Nothing
                Set oDS_PH_PY001D = Nothing
                Set oDS_PH_PY001E = Nothing
                Set oDS_PH_PY001F = Nothing
                Set oDS_PH_PY001G = Nothing
                Set oDS_PH_PY001H = Nothing
                Set oDS_PH_PY001J = Nothing
                Set oDS_PH_PY001K = Nothing
                Set oDS_PH_PY001L = Nothing
                Set oDS_PH_PY001M = Nothing
                Set oDS_PH_PY001N = Nothing
                Set oDS_PH_PY001P = Nothing
                Set oDS_PH_PY001Q = Nothing
                Set oDS_PH_PY001R = Nothing
                
                Set oMat1 = Nothing
                Set oMat2 = Nothing
                Set oMat3 = Nothing
                Set oMat4 = Nothing
                Set oMat5 = Nothing
                Set oMat6 = Nothing
                Set oMat7 = Nothing
                Set oMat8 = Nothing
                Set oMat9 = Nothing
                Set oMat10 = Nothing
                Set oMat11 = Nothing
                Set oMat12 = Nothing
                Set oMat13 = Nothing
                Set oMat14 = Nothing
                Set oMat15 = Nothing
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
            
                oForm.Items.Item("8").Width = oForm.Items.Item("Mat2").Left + oForm.Items.Item("Mat2").Width + 5
                oForm.Items.Item("8").Height = oForm.Items.Item("1").Top - 80
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
'                    Call MDC_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY001A", "Code")
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
    Sbo_Application.SetStatusBarMessage "Raise_ItemEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
                Call AuthorityCheck(oForm, "CLTCOD", "@PH_PY001A", "Code")      '//접속자 권한에 따른 사업장 보기
                
        End Select
    ElseIf (pval.BeforeAction = False) Then
        Select Case pval.MenuUID
            Case "1283":
                oForm.Mode = fm_ADD_MODE
                Call PH_PY001_FormItemEnabled
                Call PH_PY001_AddMatrixRow
            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//문서찾기
                Call PH_PY001_FormItemEnabled
                Call PH_PY001_AddMatrixRow
                oForm.Items("Code").CLICK ct_Regular
            Case "1282":    '//문서추가
                Call PH_PY001_FormItemEnabled
                Call PH_PY001_AddMatrixRow
            Case "1288", "1289", "1290", "1291":
                Call PH_PY001_FormItemEnabled
            Case "1293" '// 행삭제
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent, oMat1, oDS_PH_PY001B, "U_FILD01")       '// [MAT1] 급여 - 수당
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent, oMat2, oDS_PH_PY001C, "U_FILD01")       '// [MAT2] 급여 - 공제
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent, oMat3, oDS_PH_PY001D, "U_FamNam")       '// [MAT3] 가족사항
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent, oMat4, oDS_PH_PY001E, "U_employer")     '// [Mat4] 경력사항
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent, oMat5, oDS_PH_PY001F, "U_License")      '// [Mat5] 자격사항
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent, oMat6, oDS_PH_PY001G, "U_appNum")       '// [Mat6] 발령사항
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent, oMat7, oDS_PH_PY001H, "U_Club")         '// [Mat7] 동호회
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent, oMat8, oDS_PH_PY001J, "U_School")       '// [Mat8] 학력사항
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent, oMat9, oDS_PH_PY001K, "U_major")        '// [Mat9] 교육사항
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent, oMat10, oDS_PH_PY001L, "U_Basis")       '// [Mat10] 상벌사항
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent, oMat11, oDS_PH_PY001M, "U_Passport")    '// [Mat11] 여권사항
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent, oMat12, oDS_PH_PY001N, "U_fromDate")    '// [Mat12] 노조이력
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent, oMat13, oDS_PH_PY001P, "U_Person")      '// [Mat13] 경조사항
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent, oMat14, oDS_PH_PY001Q, "U_fromDate")    '// [Mat14] 사고현황
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent, oMat15, oDS_PH_PY001R, "U_CarNum")    '// [Mat15] 차량현황
                Call PH_PY001_AddMatrixRow
        End Select
    End If
    Call oForm.Freeze(False)
    Exit Sub
Raise_FormMenuEvent_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_FormMenuEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
                '//부서
                Set oCombo = oForm.Items("TeamCode").Specific
                If oCombo.ValidValues.Count > 0 Then
                    For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                        oCombo.ValidValues.Remove i, psk_Index
                    Next i
                End If
                
                sQry = "           SELECT       '' AS [Code], "
                sQry = sQry & "                 '' AS [Name],"
                sQry = sQry & "                 -1 AS [Seq]"
                sQry = sQry & " UNION ALL"
                sQry = sQry & " SELECT       U_Code AS [Code], "
                sQry = sQry & "                 U_CodeNm AS [Name],"
                sQry = sQry & "                 U_Seq AS [Seq]"
                sQry = sQry & " FROM        [@PS_HR200L] "
                sQry = sQry & " WHERE       Code = '1' "
                sQry = sQry & "                 AND U_Char2 = '" & Trim(oDS_PH_PY001A.GetValue("U_CLTCOD", 0)) & "'"
                sQry = sQry & "                 AND U_UseYN = 'Y'"
                sQry = sQry & " ORDER BY   Seq"
'
'
'                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
'                sQry = sQry & " WHERE Code = '1' AND U_Char2 = '" & Trim(oDS_PH_PY001A.GetValue("U_CLTCOD", 0)) & "'"
'                sQry = sQry & " ORDER BY U_Code"

                Call SetReDataCombo(oForm, sQry, oCombo)

                oForm.Items("TeamCode").DisplayDesc = True
                
        
                '//담당
                Set oCombo = oForm.Items("RspCode").Specific
                If oCombo.ValidValues.Count > 0 Then
                    For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                        oCombo.ValidValues.Remove i, psk_Index
                    Next i
                End If
                
                sQry = "           SELECT       '' AS [Code], "
                sQry = sQry & "                 '' AS [Name],"
                sQry = sQry & "                 -1 AS [Seq]"
                sQry = sQry & " UNION ALL"
                sQry = sQry & " SELECT       U_Code AS [Code], "
                sQry = sQry & "                 U_CodeNm AS [Name],"
                sQry = sQry & "                 U_Seq AS [Seq]"
                sQry = sQry & " FROM        [@PS_HR200L] "
                sQry = sQry & " WHERE       Code = '2' "
                sQry = sQry & "                 AND U_Char2 = '" & Trim(oDS_PH_PY001A.GetValue("U_CLTCOD", 0)) & "'"
                sQry = sQry & "                 AND U_Char1 = '" & Trim(oDS_PH_PY001A.GetValue("U_TeamCode", 0)) & "'"
                sQry = sQry & "                 AND U_UseYN = 'Y'"
                sQry = sQry & " ORDER BY   Seq"
                
'                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
'                sQry = sQry & " WHERE Code = '2' AND U_Char2 = '" & Trim(oDS_PH_PY001A.GetValue("U_CLTCOD", 0)) & "'"
'                sQry = sQry & " ORDER BY U_Code"

                Call SetReDataCombo(oForm, sQry, oCombo)

                oForm.Items("RspCode").DisplayDesc = True
                
        
                '//반 (사업장에 따른 반변경)
        
                Set oCombo = oForm.Items("ClsCode").Specific
        
                If oCombo.ValidValues.Count > 0 Then
                    For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                        oCombo.ValidValues.Remove i, psk_Index
                    Next i
                End If
        
                sQry = "           SELECT       '' AS [Code], "
                sQry = sQry & "                 '' AS [Name],"
                sQry = sQry & "                 -1 AS [Seq]"
                sQry = sQry & " UNION ALL"
                sQry = sQry & " SELECT       U_Code AS [Code], "
                sQry = sQry & "                 U_CodeNm AS [Name],"
                sQry = sQry & "                 U_Seq AS [Seq]"
                sQry = sQry & " FROM        [@PS_HR200L] "
                sQry = sQry & " WHERE       Code = '9' "
                sQry = sQry & "                 AND U_Char3 = '" & Trim(oDS_PH_PY001A.GetValue("U_CLTCOD", 0)) & "'"
                sQry = sQry & "                 AND U_Char2 = '" & Trim(oDS_PH_PY001A.GetValue("U_TeamCode", 0)) & "'"
                sQry = sQry & "                 AND U_Char1 = '" & Trim(oDS_PH_PY001A.GetValue("U_RspCode", 0)) & "'"
                sQry = sQry & "                 AND U_UseYN = 'Y'"
                sQry = sQry & " ORDER BY   Seq"
        
'                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
'                sQry = sQry & " WHERE Code='9' AND U_Char3 = '" & Trim(oDS_PH_PY001A.GetValue("U_CLTCOD", 0)) & "'"
'                sQry = sQry & " ORDER BY U_Code"
        
                Call SetReDataCombo(oForm, sQry, oCombo)

                oForm.Items("ClsCode").DisplayDesc = True
        
                '//근무조
                Set oCombo = oForm.Items("GNMUJO").Specific
                If oCombo.ValidValues.Count > 0 Then
                    For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                        oCombo.ValidValues.Remove i, psk_Index
                    Next i
                End If
                
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
                sQry = sQry & " WHERE Code = 'P155' AND U_Char1 = '" & Trim(oDS_PH_PY001A.GetValue("U_ShiftDat", 0)) & "'"
                sQry = sQry & " ORDER BY U_Code"

                Call SetReDataCombo(oForm, sQry, oCombo)

                oForm.Items("GNMUJO").DisplayDesc = True
                
                
                Select Case Trim(oDS_PH_PY001A.GetValue("U_PAYTYP", 0))
                Case "1"
                    
                    oForm.Items("Btn02").Visible = False
                    oForm.Items("HOBONG").Enabled = False
                    oForm.Items("YUNAMT").Enabled = True
                Case "2"
                    
                    oForm.Items("Btn02").Visible = True
                    oForm.Items("HOBONG").Enabled = True
                    oForm.Items("YUNAMT").Enabled = False
                Case "3"
                    
                    oForm.Items("Btn02").Visible = False
                    oForm.Items("HOBONG").Enabled = False
                    oForm.Items("YUNAMT").Enabled = False
                Case "4"
                    
                    oForm.Items("Btn02").Visible = False
                    oForm.Items("HOBONG").Enabled = False
                    oForm.Items("YUNAMT").Enabled = False
                End Select
                
            Case et_FORM_DATA_ADD:      '//34
            Case et_FORM_DATA_UPDATE:   '//35
                oForm.Items("Code").Enabled = False
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
    Case "Mat1", "Mat2", "Mat3", "Mat4", "Mat5", "Mat6", "Mat7", "Mat8", "Mat9", "Mat10", "Mat11", "Mat12", "Mat13", "Mat14", "Mat15"
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

Sub PH_PY001_AddMatrixRow()
    Dim oRow    As Long
    
    On Error GoTo PH_PY001_AddMatrixRow_Error
    
    Call oForm.Freeze(True)
    
    '//[Mat1]
    oMat1.FlushToDataSource
    oRow = oMat1.VisualRowCount
    
    If oMat1.VisualRowCount > 0 Then
        If Trim(oDS_PH_PY001B.GetValue("U_FILD01", oRow - 1)) <> "" Then
            If oDS_PH_PY001B.Size <= oMat1.VisualRowCount Then
                oDS_PH_PY001B.InsertRecord (oRow)
            End If
            oDS_PH_PY001B.Offset = oRow
            oDS_PH_PY001B.setValue "U_LineNum", oRow, oRow + 1
            oDS_PH_PY001B.setValue "U_FILD01", oRow, ""
            oDS_PH_PY001B.setValue "U_FILD02", oRow, ""
            oDS_PH_PY001B.setValue "U_FILD03", oRow, 0
            oMat1.LoadFromDataSource
        Else
            oDS_PH_PY001B.Offset = oRow - 1
            oDS_PH_PY001B.setValue "U_LineNum", oRow - 1, oRow
            oDS_PH_PY001B.setValue "U_FILD01", oRow - 1, ""
            oDS_PH_PY001B.setValue "U_FILD02", oRow - 1, ""
            oDS_PH_PY001B.setValue "U_FILD03", oRow - 1, 0
            oMat1.LoadFromDataSource
        End If
    ElseIf oMat1.VisualRowCount = 0 Then
        oDS_PH_PY001B.Offset = oRow
        oDS_PH_PY001B.setValue "U_LineNum", oRow, oRow + 1
        oDS_PH_PY001B.setValue "U_FILD01", oRow, ""
        oDS_PH_PY001B.setValue "U_FILD02", oRow, ""
        oDS_PH_PY001B.setValue "U_FILD03", oRow, 0
        oMat1.LoadFromDataSource
    End If
    
    '//[Mat2]
    oMat2.FlushToDataSource
    oRow = oMat2.VisualRowCount
    
    If oMat2.VisualRowCount > 0 Then
        If Trim(oDS_PH_PY001C.GetValue("U_FILD01", oRow - 1)) <> "" Then
            If oDS_PH_PY001C.Size <= oMat2.VisualRowCount Then
                oDS_PH_PY001C.InsertRecord (oRow)
            End If
            oDS_PH_PY001C.Offset = oRow
            oDS_PH_PY001C.setValue "U_LineNum", oRow, oRow + 1
            oDS_PH_PY001C.setValue "U_FILD01", oRow, ""
            oDS_PH_PY001C.setValue "U_FILD02", oRow, ""
            oDS_PH_PY001C.setValue "U_FILD03", oRow, 0
            oMat2.LoadFromDataSource
        Else
            oDS_PH_PY001C.Offset = oRow - 1
            oDS_PH_PY001C.setValue "U_LineNum", oRow - 1, oRow
            oDS_PH_PY001C.setValue "U_FILD01", oRow - 1, ""
            oDS_PH_PY001C.setValue "U_FILD02", oRow - 1, ""
            oDS_PH_PY001C.setValue "U_FILD03", oRow - 1, 0
            oMat2.LoadFromDataSource
        End If
    ElseIf oMat2.VisualRowCount = 0 Then
        oDS_PH_PY001C.Offset = oRow
        oDS_PH_PY001C.setValue "U_LineNum", oRow, oRow + 1
        oDS_PH_PY001C.setValue "U_FILD01", oRow, ""
        oDS_PH_PY001C.setValue "U_FILD02", oRow, ""
        oDS_PH_PY001C.setValue "U_FILD03", oRow, 0
        oMat2.LoadFromDataSource
    End If
    
    '//[Mat3]
    oMat3.FlushToDataSource
    oRow = oMat3.VisualRowCount
    
    If oMat3.VisualRowCount > 0 Then
        If Trim(oDS_PH_PY001D.GetValue("U_FamNam", oRow - 1)) <> "" Then
            If oDS_PH_PY001D.Size <= oMat3.VisualRowCount Then
                oDS_PH_PY001D.InsertRecord (oRow)
            End If
            oDS_PH_PY001D.Offset = oRow
            oDS_PH_PY001D.setValue "U_LineNum", oRow, oRow + 1
            oDS_PH_PY001D.setValue "U_FamNam", oRow, ""
            oDS_PH_PY001D.setValue "U_FamPer", oRow, ""
            oDS_PH_PY001D.setValue "U_BIRGBN", oRow, ""
            oDS_PH_PY001D.setValue "U_BIRDAT", oRow, ""
            oDS_PH_PY001D.setValue "U_FamSch", oRow, ""
            oDS_PH_PY001D.setValue "U_FamJob", oRow, ""
            oDS_PH_PY001D.setValue "U_FamGun", oRow, ""
            oDS_PH_PY001D.setValue "U_JsnType", oRow, ""
            oDS_PH_PY001D.setValue "U_ChkCod", oRow, ""
            oDS_PH_PY001D.setValue "U_ChkInt", oRow, ""
            oDS_PH_PY001D.setValue "U_CHKTAX", oRow, "N"
            oDS_PH_PY001D.setValue "U_ChkBas", oRow, "N"
            oDS_PH_PY001D.setValue "U_ChkJan", oRow, "N"
            oDS_PH_PY001D.setValue "U_ChkBoH", oRow, "N"
            oDS_PH_PY001D.setValue "U_ChkChl", oRow, "N"
            oDS_PH_PY001D.setValue "U_ChkEdu", oRow, "N"
            oDS_PH_PY001D.setValue "U_Remark", oRow, ""
            oMat3.LoadFromDataSource
        Else
            oDS_PH_PY001D.Offset = oRow - 1
            oDS_PH_PY001D.setValue "U_LineNum", oRow - 1, oRow
            oDS_PH_PY001D.setValue "U_FamNam", oRow - 1, ""
            oDS_PH_PY001D.setValue "U_FamPer", oRow - 1, ""
            oDS_PH_PY001D.setValue "U_BIRGBN", oRow - 1, ""
            oDS_PH_PY001D.setValue "U_BIRDAT", oRow - 1, ""
            oDS_PH_PY001D.setValue "U_FamSch", oRow - 1, ""
            oDS_PH_PY001D.setValue "U_FamJob", oRow - 1, ""
            oDS_PH_PY001D.setValue "U_FamGun", oRow - 1, ""
            oDS_PH_PY001D.setValue "U_JsnType", oRow - 1, ""
            oDS_PH_PY001D.setValue "U_ChkCod", oRow - 1, ""
            oDS_PH_PY001D.setValue "U_ChkInt", oRow - 1, ""
            oDS_PH_PY001D.setValue "U_CHKTAX", oRow - 1, "N"
            oDS_PH_PY001D.setValue "U_ChkBas", oRow - 1, "N"
            oDS_PH_PY001D.setValue "U_ChkJan", oRow - 1, "N"
            oDS_PH_PY001D.setValue "U_ChkBoH", oRow - 1, "N"
            oDS_PH_PY001D.setValue "U_ChkChl", oRow - 1, "N"
            oDS_PH_PY001D.setValue "U_ChkEdu", oRow - 1, "N"
            oDS_PH_PY001D.setValue "U_Remark", oRow - 1, ""
            oMat3.LoadFromDataSource
        End If
    ElseIf oMat3.VisualRowCount = 0 Then
        oDS_PH_PY001D.Offset = oRow
        oDS_PH_PY001D.setValue "U_LineNum", oRow, oRow + 1
        oDS_PH_PY001D.setValue "U_FamNam", oRow, ""
        oDS_PH_PY001D.setValue "U_FamPer", oRow, ""
        oDS_PH_PY001D.setValue "U_BIRGBN", oRow, ""
        oDS_PH_PY001D.setValue "U_BIRDAT", oRow, ""
        oDS_PH_PY001D.setValue "U_FamSch", oRow, ""
        oDS_PH_PY001D.setValue "U_FamJob", oRow, ""
        oDS_PH_PY001D.setValue "U_FamGun", oRow, ""
        oDS_PH_PY001D.setValue "U_JsnType", oRow, ""
        oDS_PH_PY001D.setValue "U_ChkCod", oRow, ""
        oDS_PH_PY001D.setValue "U_ChkInt", oRow, ""
        oDS_PH_PY001D.setValue "U_CHKTAX", oRow, "N"
        oDS_PH_PY001D.setValue "U_ChkBas", oRow, "N"
        oDS_PH_PY001D.setValue "U_ChkJan", oRow, "N"
        oDS_PH_PY001D.setValue "U_ChkBoH", oRow, "N"
        oDS_PH_PY001D.setValue "U_ChkChl", oRow, "N"
        oDS_PH_PY001D.setValue "U_ChkEdu", oRow, "N"
        oDS_PH_PY001D.setValue "U_Remark", oRow, ""
        oMat3.LoadFromDataSource
    End If
    
    '//[Mat4]
    oMat4.FlushToDataSource
    oRow = oMat4.VisualRowCount
    
    If oMat4.VisualRowCount > 0 Then
        If Trim(oDS_PH_PY001E.GetValue("U_employer", oRow - 1)) <> "" Then
            If oDS_PH_PY001E.Size <= oMat4.VisualRowCount Then
                oDS_PH_PY001E.InsertRecord (oRow)
            End If
            oDS_PH_PY001E.Offset = oRow
            oDS_PH_PY001E.setValue "U_LineNum", oRow, oRow + 1
            oDS_PH_PY001E.setValue "U_fromDate", oRow, ""
            oDS_PH_PY001E.setValue "U_toDate", oRow, ""
            oDS_PH_PY001E.setValue "U_employer", oRow, ""
            oDS_PH_PY001E.setValue "U_depart", oRow, ""
            oDS_PH_PY001E.setValue "U_position", oRow, ""
            oDS_PH_PY001E.setValue "U_remarks", oRow, ""
            oMat4.LoadFromDataSource
        Else
            oDS_PH_PY001E.Offset = oRow - 1
            oDS_PH_PY001E.setValue "U_LineNum", oRow - 1, oRow
            oDS_PH_PY001E.setValue "U_fromDate", oRow - 1, ""
            oDS_PH_PY001E.setValue "U_toDate", oRow - 1, ""
            oDS_PH_PY001E.setValue "U_employer", oRow - 1, ""
            oDS_PH_PY001E.setValue "U_depart", oRow - 1, ""
            oDS_PH_PY001E.setValue "U_position", oRow - 1, ""
            oDS_PH_PY001E.setValue "U_remarks", oRow - 1, ""
            oMat4.LoadFromDataSource
        End If
    ElseIf oMat4.VisualRowCount = 0 Then
        oDS_PH_PY001E.Offset = oRow
        oDS_PH_PY001E.setValue "U_LineNum", oRow, oRow + 1
        oDS_PH_PY001E.setValue "U_fromDate", oRow, ""
        oDS_PH_PY001E.setValue "U_toDate", oRow, ""
        oDS_PH_PY001E.setValue "U_employer", oRow, ""
        oDS_PH_PY001E.setValue "U_depart", oRow, ""
        oDS_PH_PY001E.setValue "U_position", oRow, ""
        oDS_PH_PY001E.setValue "U_remarks", oRow, ""
        oMat4.LoadFromDataSource
    End If
    
    '//[Mat5]
    oMat5.FlushToDataSource
    oRow = oMat5.VisualRowCount
    
    If oMat5.VisualRowCount > 0 Then
        If Trim(oDS_PH_PY001F.GetValue("U_License", oRow - 1)) <> "" Then
            If oDS_PH_PY001F.Size <= oMat5.VisualRowCount Then
                oDS_PH_PY001F.InsertRecord (oRow)
            End If
            oDS_PH_PY001F.Offset = oRow
            oDS_PH_PY001F.setValue "U_LineNum", oRow, oRow + 1
            oDS_PH_PY001F.setValue "U_License", oRow, ""
            oDS_PH_PY001F.setValue "U_Grade", oRow, ""
            oDS_PH_PY001F.setValue "U_LCNumber", oRow, ""
            oDS_PH_PY001F.setValue "U_IssuInst", oRow, ""
            oDS_PH_PY001F.setValue "U_acquDate", oRow, ""
            oDS_PH_PY001F.setValue "U_loseDate", oRow, ""
            oDS_PH_PY001F.setValue "U_Require", oRow, ""
            oDS_PH_PY001F.setValue "U_LCPay", oRow, ""
            oDS_PH_PY001F.setValue "U_PayDate", oRow, ""
            oMat5.LoadFromDataSource
        Else
            oDS_PH_PY001F.Offset = oRow - 1
            oDS_PH_PY001F.setValue "U_LineNum", oRow - 1, oRow
            oDS_PH_PY001F.setValue "U_License", oRow - 1, ""
            oDS_PH_PY001F.setValue "U_Grade", oRow - 1, ""
            oDS_PH_PY001F.setValue "U_LCNumber", oRow - 1, ""
            oDS_PH_PY001F.setValue "U_IssuInst", oRow - 1, ""
            oDS_PH_PY001F.setValue "U_acquDate", oRow - 1, ""
            oDS_PH_PY001F.setValue "U_loseDate", oRow - 1, ""
            oDS_PH_PY001F.setValue "U_Require", oRow - 1, ""
            oDS_PH_PY001F.setValue "U_LCPay", oRow - 1, ""
            oDS_PH_PY001F.setValue "U_PayDate", oRow - 1, ""
            oMat5.LoadFromDataSource
        End If
    ElseIf oMat5.VisualRowCount = 0 Then
        oDS_PH_PY001F.Offset = oRow
        oDS_PH_PY001F.setValue "U_LineNum", oRow, oRow + 1
        oDS_PH_PY001F.setValue "U_License", oRow, ""
        oDS_PH_PY001F.setValue "U_Grade", oRow, ""
        oDS_PH_PY001F.setValue "U_LCNumber", oRow, ""
        oDS_PH_PY001F.setValue "U_IssuInst", oRow, ""
        oDS_PH_PY001F.setValue "U_acquDate", oRow, ""
        oDS_PH_PY001F.setValue "U_loseDate", oRow, ""
        oDS_PH_PY001F.setValue "U_Require", oRow, ""
        oDS_PH_PY001F.setValue "U_LCPay", oRow, ""
        oDS_PH_PY001F.setValue "U_PayDate", oRow, ""
        oMat5.LoadFromDataSource
    End If
    
    '//[Mat6]
    oMat6.FlushToDataSource
    oRow = oMat6.VisualRowCount
    
    If oMat6.VisualRowCount > 0 Then
        If Trim(oDS_PH_PY001G.GetValue("U_appNum", oRow - 1)) <> "" Then
            If oDS_PH_PY001G.Size <= oMat6.VisualRowCount Then
                oDS_PH_PY001G.InsertRecord (oRow)
            End If
            oDS_PH_PY001G.Offset = oRow
            oDS_PH_PY001G.setValue "U_LineNum", oRow, oRow + 1
            oDS_PH_PY001G.setValue "U_appDate", oRow, ""
            oDS_PH_PY001G.setValue "U_appType", oRow, ""
            oDS_PH_PY001G.setValue "U_appNum", oRow, ""
            oDS_PH_PY001G.setValue "U_TeamCode", oRow, ""
            oDS_PH_PY001G.setValue "U_RspCode", oRow, ""
            oDS_PH_PY001G.setValue "U_JIGTYP", oRow, ""
            oDS_PH_PY001G.setValue "U_position", oRow, ""
            oDS_PH_PY001G.setValue "U_CallName", oRow, ""
            oDS_PH_PY001G.setValue "U_JIGCOD", oRow, ""
            oDS_PH_PY001G.setValue "U_HOBONG", oRow, ""
            oDS_PH_PY001G.setValue "U_STDAMT", oRow, "0"
            oMat6.LoadFromDataSource
        Else
            oDS_PH_PY001G.Offset = oRow - 1
            oDS_PH_PY001G.setValue "U_LineNum", oRow - 1, oRow
            oDS_PH_PY001G.setValue "U_appDate", oRow - 1, ""
            oDS_PH_PY001G.setValue "U_appType", oRow - 1, ""
            oDS_PH_PY001G.setValue "U_appNum", oRow - 1, ""
            oDS_PH_PY001G.setValue "U_TeamCode", oRow - 1, ""
            oDS_PH_PY001G.setValue "U_RspCode", oRow - 1, ""
            oDS_PH_PY001G.setValue "U_JIGTYP", oRow - 1, ""
            oDS_PH_PY001G.setValue "U_position", oRow - 1, ""
            oDS_PH_PY001G.setValue "U_CallName", oRow - 1, ""
            oDS_PH_PY001G.setValue "U_JIGCOD", oRow - 1, ""
            oDS_PH_PY001G.setValue "U_HOBONG", oRow - 1, ""
            oDS_PH_PY001G.setValue "U_STDAMT", oRow - 1, "0"
            oMat6.LoadFromDataSource
        End If
    ElseIf oMat6.VisualRowCount = 0 Then
        oDS_PH_PY001G.Offset = oRow
        oDS_PH_PY001G.setValue "U_LineNum", oRow, oRow + 1
        oDS_PH_PY001G.setValue "U_appDate", oRow, ""
            oDS_PH_PY001G.setValue "U_appType", oRow, ""
            oDS_PH_PY001G.setValue "U_appNum", oRow, ""
            oDS_PH_PY001G.setValue "U_TeamCode", oRow, ""
            oDS_PH_PY001G.setValue "U_RspCode", oRow, ""
            oDS_PH_PY001G.setValue "U_JIGTYP", oRow, ""
            oDS_PH_PY001G.setValue "U_position", oRow, ""
            oDS_PH_PY001G.setValue "U_CallName", oRow, ""
            oDS_PH_PY001G.setValue "U_JIGCOD", oRow, ""
            oDS_PH_PY001G.setValue "U_HOBONG", oRow, ""
            oDS_PH_PY001G.setValue "U_STDAMT", oRow, "0"
        oMat6.LoadFromDataSource
    End If
    '//[Mat7]
    oMat7.FlushToDataSource
    oRow = oMat7.VisualRowCount
    
    If oMat7.VisualRowCount > 0 Then
        If Trim(oDS_PH_PY001H.GetValue("U_InDate", oRow - 1)) <> "" Then
            If oDS_PH_PY001H.Size <= oMat7.VisualRowCount Then
                oDS_PH_PY001H.InsertRecord (oRow)
            End If
            oDS_PH_PY001H.Offset = oRow
            oDS_PH_PY001H.setValue "U_LineNum", oRow, oRow + 1
            oDS_PH_PY001H.setValue "U_Club", oRow, ""
            oDS_PH_PY001H.setValue "U_InDate", oRow, ""
            oDS_PH_PY001H.setValue "U_OutDate", oRow, ""
            oDS_PH_PY001H.setValue "U_Position", oRow, ""
            oDS_PH_PY001H.setValue "U_Comments", oRow, ""
            oMat7.LoadFromDataSource
        Else
            oDS_PH_PY001H.Offset = oRow - 1
            oDS_PH_PY001H.setValue "U_LineNum", oRow - 1, oRow
            oDS_PH_PY001H.setValue "U_Club", oRow - 1, ""
            oDS_PH_PY001H.setValue "U_InDate", oRow - 1, ""
            oDS_PH_PY001H.setValue "U_OutDate", oRow - 1, ""
            oDS_PH_PY001H.setValue "U_Position", oRow - 1, ""
            oDS_PH_PY001H.setValue "U_Comments", oRow - 1, ""
            oMat7.LoadFromDataSource
        End If
    ElseIf oMat7.VisualRowCount = 0 Then
        oDS_PH_PY001H.Offset = oRow
        oDS_PH_PY001H.setValue "U_LineNum", oRow, oRow + 1
        oDS_PH_PY001H.setValue "U_Club", oRow, ""
            oDS_PH_PY001H.setValue "U_InDate", oRow, ""
            oDS_PH_PY001H.setValue "U_OutDate", oRow, ""
            oDS_PH_PY001H.setValue "U_Position", oRow, ""
            oDS_PH_PY001H.setValue "U_Comments", oRow, ""
        oMat7.LoadFromDataSource
    End If
    '//[Mat8]
    oMat8.FlushToDataSource
    oRow = oMat8.VisualRowCount
    
    If oMat8.VisualRowCount > 0 Then
        If Trim(oDS_PH_PY001J.GetValue("U_School", oRow - 1)) <> "" Then
            If oDS_PH_PY001J.Size <= oMat8.VisualRowCount Then
                oDS_PH_PY001J.InsertRecord (oRow)
            End If
            oDS_PH_PY001J.Offset = oRow
            oDS_PH_PY001J.setValue "U_LineNum", oRow, oRow + 1
            oDS_PH_PY001J.setValue "U_fromDate", oRow, ""
            oDS_PH_PY001J.setValue "U_toDate", oRow, ""
            oDS_PH_PY001J.setValue "U_School", oRow, ""
            oDS_PH_PY001J.setValue "U_depart", oRow, ""
            oDS_PH_PY001J.setValue "U_degree", oRow, ""
            oDS_PH_PY001J.setValue "U_complete", oRow, ""
            oMat8.LoadFromDataSource
        Else
            oDS_PH_PY001J.Offset = oRow - 1
            oDS_PH_PY001J.setValue "U_LineNum", oRow - 1, oRow
            oDS_PH_PY001J.setValue "U_fromDate", oRow - 1, ""
            oDS_PH_PY001J.setValue "U_toDate", oRow - 1, ""
            oDS_PH_PY001J.setValue "U_School", oRow - 1, ""
            oDS_PH_PY001J.setValue "U_depart", oRow - 1, ""
            oDS_PH_PY001J.setValue "U_degree", oRow - 1, ""
            oDS_PH_PY001J.setValue "U_complete", oRow - 1, ""
            oMat8.LoadFromDataSource
        End If
    ElseIf oMat8.VisualRowCount = 0 Then
        oDS_PH_PY001J.Offset = oRow
        oDS_PH_PY001J.setValue "U_LineNum", oRow, oRow + 1
        oDS_PH_PY001J.setValue "U_fromDate", oRow, ""
        oDS_PH_PY001J.setValue "U_toDate", oRow, ""
        oDS_PH_PY001J.setValue "U_School", oRow, ""
        oDS_PH_PY001J.setValue "U_depart", oRow, ""
        oDS_PH_PY001J.setValue "U_degree", oRow, ""
        oDS_PH_PY001J.setValue "U_complete", oRow, ""
        oMat8.LoadFromDataSource
    End If
    '//[Mat9]
    oMat9.FlushToDataSource
    oRow = oMat9.VisualRowCount
    
    If oMat9.VisualRowCount > 0 Then
        If Trim(oDS_PH_PY001K.GetValue("U_major", oRow - 1)) <> "" Then
            If oDS_PH_PY001K.Size <= oMat9.VisualRowCount Then
                oDS_PH_PY001K.InsertRecord (oRow)
            End If
            oDS_PH_PY001K.Offset = oRow
            oDS_PH_PY001K.setValue "U_LineNum", oRow, oRow + 1
            oDS_PH_PY001K.setValue "U_type", oRow, ""
            oDS_PH_PY001K.setValue "U_major", oRow, ""
            oDS_PH_PY001K.setValue "U_fromDate", oRow, ""
            oDS_PH_PY001K.setValue "U_toDate", oRow, ""
            oDS_PH_PY001K.setValue "U_institut", oRow, ""
            oDS_PH_PY001K.setValue "U_diploma", oRow, ""
            oDS_PH_PY001K.setValue "U_EduExp", oRow, ""
            oDS_PH_PY001K.setValue "U_TraExp", oRow, ""
            oMat9.LoadFromDataSource
        Else
            oDS_PH_PY001K.Offset = oRow - 1
            oDS_PH_PY001K.setValue "U_LineNum", oRow - 1, oRow
            oDS_PH_PY001K.setValue "U_type", oRow - 1, ""
            oDS_PH_PY001K.setValue "U_major", oRow - 1, ""
            oDS_PH_PY001K.setValue "U_fromDate", oRow - 1, ""
            oDS_PH_PY001K.setValue "U_toDate", oRow - 1, ""
            oDS_PH_PY001K.setValue "U_institut", oRow - 1, ""
            oDS_PH_PY001K.setValue "U_diploma", oRow - 1, ""
            oDS_PH_PY001K.setValue "U_EduExp", oRow - 1, ""
            oDS_PH_PY001K.setValue "U_TraExp", oRow - 1, ""
            oMat9.LoadFromDataSource
        End If
    ElseIf oMat9.VisualRowCount = 0 Then
        oDS_PH_PY001K.Offset = oRow
        oDS_PH_PY001K.setValue "U_LineNum", oRow, oRow + 1
        oDS_PH_PY001K.setValue "U_type", oRow, ""
        oDS_PH_PY001K.setValue "U_major", oRow, ""
        oDS_PH_PY001K.setValue "U_fromDate", oRow, ""
        oDS_PH_PY001K.setValue "U_toDate", oRow, ""
        oDS_PH_PY001K.setValue "U_institut", oRow, ""
        oDS_PH_PY001K.setValue "U_diploma", oRow, ""
        oDS_PH_PY001K.setValue "U_EduExp", oRow, ""
        oDS_PH_PY001K.setValue "U_TraExp", oRow, ""
        oMat9.LoadFromDataSource
    End If
    '//[Mat10]
    oMat10.FlushToDataSource
    oRow = oMat10.VisualRowCount
    
    If oMat10.VisualRowCount > 0 Then
        If Trim(oDS_PH_PY001L.GetValue("U_Basis", oRow - 1)) <> "" Then
            If oDS_PH_PY001L.Size <= oMat10.VisualRowCount Then
                oDS_PH_PY001L.InsertRecord (oRow)
            End If
            oDS_PH_PY001L.Offset = oRow
            oDS_PH_PY001L.setValue "U_LineNum", oRow, oRow + 1
            oDS_PH_PY001L.setValue "U_Date", oRow, ""
            oDS_PH_PY001L.setValue "U_Category", oRow, ""
            oDS_PH_PY001L.setValue "U_Type", oRow, ""
            oDS_PH_PY001L.setValue "U_Basis", oRow, ""
            oDS_PH_PY001L.setValue "U_Text", oRow, ""
            oDS_PH_PY001L.setValue "U_Reward", oRow, ""
            oDS_PH_PY001L.setValue "U_Reduce", oRow, ""
            oDS_PH_PY001L.setValue "U_addition", oRow, ""
            oDS_PH_PY001L.setValue "U_fromYM", oRow, ""
            oDS_PH_PY001L.setValue "U_toYM", oRow, ""
            oMat10.LoadFromDataSource
        Else
            oDS_PH_PY001L.Offset = oRow - 1
            oDS_PH_PY001L.setValue "U_LineNum", oRow - 1, oRow
            oDS_PH_PY001L.setValue "U_Date", oRow - 1, ""
            oDS_PH_PY001L.setValue "U_Category", oRow - 1, ""
            oDS_PH_PY001L.setValue "U_Type", oRow - 1, ""
            oDS_PH_PY001L.setValue "U_Basis", oRow - 1, ""
            oDS_PH_PY001L.setValue "U_Text", oRow - 1, ""
            oDS_PH_PY001L.setValue "U_Reward", oRow - 1, ""
            oDS_PH_PY001L.setValue "U_Reduce", oRow - 1, ""
            oDS_PH_PY001L.setValue "U_addition", oRow - 1, ""
            oDS_PH_PY001L.setValue "U_fromYM", oRow - 1, ""
            oDS_PH_PY001L.setValue "U_toYM", oRow - 1, ""
            oMat10.LoadFromDataSource
        End If
    ElseIf oMat10.VisualRowCount = 0 Then
        oDS_PH_PY001L.Offset = oRow
        oDS_PH_PY001L.setValue "U_LineNum", oRow, oRow + 1
        oDS_PH_PY001L.setValue "U_Date", oRow, ""
        oDS_PH_PY001L.setValue "U_Category", oRow, ""
        oDS_PH_PY001L.setValue "U_Type", oRow, ""
        oDS_PH_PY001L.setValue "U_Basis", oRow, ""
        oDS_PH_PY001L.setValue "U_Text", oRow, ""
        oDS_PH_PY001L.setValue "U_Reward", oRow, ""
        oDS_PH_PY001L.setValue "U_Reduce", oRow, ""
        oDS_PH_PY001L.setValue "U_addition", oRow, ""
        oDS_PH_PY001L.setValue "U_fromYM", oRow, ""
        oDS_PH_PY001L.setValue "U_toYM", oRow, ""
        oMat10.LoadFromDataSource
    End If
    '//[Mat11]
    oMat11.FlushToDataSource
    oRow = oMat11.VisualRowCount
    
    If oMat11.VisualRowCount > 0 Then
        If Trim(oDS_PH_PY001M.GetValue("U_Passport", oRow - 1)) <> "" Then
            If oDS_PH_PY001M.Size <= oMat11.VisualRowCount Then
                oDS_PH_PY001M.InsertRecord (oRow)
            End If
            oDS_PH_PY001M.Offset = oRow
            oDS_PH_PY001M.setValue "U_LineNum", oRow, oRow + 1
            oDS_PH_PY001M.setValue "U_Passport", oRow, ""
            oDS_PH_PY001M.setValue "U_IssueDat", oRow, ""
            oDS_PH_PY001M.setValue "U_ExpirDat", oRow, ""
            oDS_PH_PY001M.setValue "U_PassType", oRow, ""
            oDS_PH_PY001M.setValue "U_Comments", oRow, ""
            oMat11.LoadFromDataSource
        Else
            oDS_PH_PY001M.Offset = oRow - 1
            oDS_PH_PY001M.setValue "U_LineNum", oRow - 1, oRow
            oDS_PH_PY001M.setValue "U_Passport", oRow - 1, ""
            oDS_PH_PY001M.setValue "U_IssueDat", oRow - 1, ""
            oDS_PH_PY001M.setValue "U_ExpirDat", oRow - 1, ""
            oDS_PH_PY001M.setValue "U_PassType", oRow - 1, ""
            oDS_PH_PY001M.setValue "U_Comments", oRow - 1, ""
            oMat11.LoadFromDataSource
        End If
    ElseIf oMat11.VisualRowCount = 0 Then
        oDS_PH_PY001M.Offset = oRow
        oDS_PH_PY001M.setValue "U_LineNum", oRow, oRow + 1
        oDS_PH_PY001M.setValue "U_Passport", oRow, ""
        oDS_PH_PY001M.setValue "U_IssueDat", oRow, ""
        oDS_PH_PY001M.setValue "U_ExpirDat", oRow, ""
        oDS_PH_PY001M.setValue "U_PassType", oRow, ""
        oDS_PH_PY001M.setValue "U_Comments", oRow, ""
        
        oMat11.LoadFromDataSource
    End If
    '//[Mat12]
    oMat12.FlushToDataSource
    oRow = oMat12.VisualRowCount
    
    If oMat12.VisualRowCount > 0 Then
        If Trim(oDS_PH_PY001N.GetValue("U_Period", oRow - 1)) <> "" Then
            If oDS_PH_PY001N.Size <= oMat12.VisualRowCount Then
                oDS_PH_PY001N.InsertRecord (oRow)
            End If
            oDS_PH_PY001N.Offset = oRow
            oDS_PH_PY001N.setValue "U_LineNum", oRow, oRow + 1
            oDS_PH_PY001N.setValue "U_fromDate", oRow, ""
            oDS_PH_PY001N.setValue "U_toDate", oRow, ""
            oDS_PH_PY001N.setValue "U_Period", oRow, ""
            oDS_PH_PY001N.setValue "U_Position", oRow, ""
            oDS_PH_PY001N.setValue "U_Comments", oRow, ""
            oMat12.LoadFromDataSource
        Else
            oDS_PH_PY001N.Offset = oRow - 1
            oDS_PH_PY001N.setValue "U_LineNum", oRow - 1, oRow
            oDS_PH_PY001N.setValue "U_fromDate", oRow - 1, ""
            oDS_PH_PY001N.setValue "U_toDate", oRow - 1, ""
            oDS_PH_PY001N.setValue "U_Period", oRow - 1, ""
            oDS_PH_PY001N.setValue "U_Position", oRow - 1, ""
            oDS_PH_PY001N.setValue "U_Comments", oRow - 1, ""
            oMat12.LoadFromDataSource
        End If
    ElseIf oMat12.VisualRowCount = 0 Then
        oDS_PH_PY001N.Offset = oRow
        oDS_PH_PY001N.setValue "U_LineNum", oRow, oRow + 1
        oDS_PH_PY001N.setValue "U_fromDate", oRow, ""
        oDS_PH_PY001N.setValue "U_toDate", oRow, ""
        oDS_PH_PY001N.setValue "U_Period", oRow, ""
        oDS_PH_PY001N.setValue "U_Position", oRow, ""
        oDS_PH_PY001N.setValue "U_Comments", oRow, ""
        oMat12.LoadFromDataSource
    End If
    '//[Mat13]
    oMat13.FlushToDataSource
    oRow = oMat13.VisualRowCount
    
    If oMat13.VisualRowCount > 0 Then
        If Trim(oDS_PH_PY001P.GetValue("U_Date", oRow - 1)) <> "" Then
            If oDS_PH_PY001P.Size <= oMat13.VisualRowCount Then
                oDS_PH_PY001P.InsertRecord (oRow)
            End If
            oDS_PH_PY001P.Offset = oRow
            oDS_PH_PY001P.setValue "U_LineNum", oRow, oRow + 1
            oDS_PH_PY001P.setValue "U_Date", oRow, ""
            oDS_PH_PY001P.setValue "U_Relation", oRow, ""
            oDS_PH_PY001P.setValue "U_Person", oRow, ""
            oDS_PH_PY001P.setValue "U_ComExpen", oRow, ""
            oDS_PH_PY001P.setValue "U_empExpen", oRow, ""
            oDS_PH_PY001P.setValue "U_Comments", oRow, ""
            oMat13.LoadFromDataSource
        Else
            oDS_PH_PY001P.Offset = oRow - 1
            oDS_PH_PY001P.setValue "U_LineNum", oRow - 1, oRow
            oDS_PH_PY001P.setValue "U_Date", oRow - 1, ""
            oDS_PH_PY001P.setValue "U_Relation", oRow - 1, ""
            oDS_PH_PY001P.setValue "U_Person", oRow - 1, ""
            oDS_PH_PY001P.setValue "U_ComExpen", oRow - 1, ""
            oDS_PH_PY001P.setValue "U_empExpen", oRow - 1, ""
            oDS_PH_PY001P.setValue "U_Comments", oRow - 1, ""
            oMat13.LoadFromDataSource
        End If
    ElseIf oMat13.VisualRowCount = 0 Then
        oDS_PH_PY001P.Offset = oRow
        oDS_PH_PY001P.setValue "U_LineNum", oRow, oRow + 1
        oDS_PH_PY001P.setValue "U_Date", oRow, ""
        oDS_PH_PY001P.setValue "U_Relation", oRow, ""
        oDS_PH_PY001P.setValue "U_Person", oRow, ""
        oDS_PH_PY001P.setValue "U_ComExpen", oRow, ""
        oDS_PH_PY001P.setValue "U_empExpen", oRow, ""
        oDS_PH_PY001P.setValue "U_Comments", oRow, ""
        oMat13.LoadFromDataSource
    End If
    '//[Mat14]
    oMat14.FlushToDataSource
    oRow = oMat14.VisualRowCount
    
    If oMat14.VisualRowCount > 0 Then
        If Trim(oDS_PH_PY001Q.GetValue("U_injury", oRow - 1)) <> "" Then
            If oDS_PH_PY001Q.Size <= oMat14.VisualRowCount Then
                oDS_PH_PY001Q.InsertRecord (oRow)
            End If
            oDS_PH_PY001Q.Offset = oRow
            oDS_PH_PY001Q.setValue "U_LineNum", oRow, oRow + 1
            oDS_PH_PY001Q.setValue "U_type", oRow, ""
            oDS_PH_PY001Q.setValue "U_fromDate", oRow, ""
            oDS_PH_PY001Q.setValue "U_toDate", oRow, ""
            oDS_PH_PY001Q.setValue "U_injury", oRow, ""
            oDS_PH_PY001Q.setValue "U_details", oRow, ""
            oMat14.LoadFromDataSource
        Else
            oDS_PH_PY001Q.Offset = oRow - 1
            oDS_PH_PY001Q.setValue "U_LineNum", oRow - 1, oRow
            oDS_PH_PY001Q.setValue "U_type", oRow - 1, ""
            oDS_PH_PY001Q.setValue "U_fromDate", oRow - 1, ""
            oDS_PH_PY001Q.setValue "U_toDate", oRow - 1, ""
            oDS_PH_PY001Q.setValue "U_injury", oRow - 1, ""
            oDS_PH_PY001Q.setValue "U_details", oRow - 1, ""
            oMat14.LoadFromDataSource
        End If
    ElseIf oMat14.VisualRowCount = 0 Then
        oDS_PH_PY001Q.Offset = oRow
        oDS_PH_PY001Q.setValue "U_LineNum", oRow, oRow + 1
        oDS_PH_PY001Q.setValue "U_type", oRow, ""
        oDS_PH_PY001Q.setValue "U_fromDate", oRow, ""
        oDS_PH_PY001Q.setValue "U_toDate", oRow, ""
        oDS_PH_PY001Q.setValue "U_injury", oRow, ""
        oDS_PH_PY001Q.setValue "U_details", oRow, ""
        oMat14.LoadFromDataSource
    End If
    
    
   '//[Mat15]
    oMat15.FlushToDataSource
    oRow = oMat15.VisualRowCount
    
    If oMat15.VisualRowCount > 0 Then
        If Trim(oDS_PH_PY001R.GetValue("U_CarNum", oRow - 1)) <> "" Then
            If oDS_PH_PY001R.Size <= oMat15.VisualRowCount Then
                oDS_PH_PY001R.InsertRecord (oRow)
            End If
            oDS_PH_PY001R.Offset = oRow
            oDS_PH_PY001R.setValue "U_LineNum", oRow, oRow + 1
            oDS_PH_PY001R.setValue "U_CarType", oRow, ""
            oDS_PH_PY001R.setValue "U_CarNum", oRow, ""
            oDS_PH_PY001R.setValue "U_InDate", oRow, ""
            oDS_PH_PY001R.setValue "U_ETC12", oRow, ""
            oMat15.LoadFromDataSource
        Else
            oDS_PH_PY001R.Offset = oRow - 1
            oDS_PH_PY001R.setValue "U_LineNum", oRow, oRow + 1
            oDS_PH_PY001R.setValue "U_CarType", oRow, ""
            oDS_PH_PY001R.setValue "U_CarNum", oRow, ""
            oDS_PH_PY001R.setValue "U_InDate", oRow, ""
            oDS_PH_PY001R.setValue "U_ETC12", oRow, ""
            oMat15.LoadFromDataSource
        End If
    ElseIf oMat15.VisualRowCount = 0 Then
            oDS_PH_PY001R.Offset = oRow
            oDS_PH_PY001R.setValue "U_LineNum", oRow, oRow + 1
            oDS_PH_PY001R.setValue "U_CarType", oRow, ""
            oDS_PH_PY001R.setValue "U_CarNum", oRow, ""
            oDS_PH_PY001R.setValue "U_InDate", oRow, ""
            oDS_PH_PY001R.setValue "U_ETC12", oRow, ""
    End If

    Call oForm.Freeze(False)
    Exit Sub
PH_PY001_AddMatrixRow_Error:

    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY001_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY001_FormClear()
On Error GoTo PH_PY001_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY001'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm.Items("DocEntry").Specific.VALUE = DocEntry
    End If
    Exit Sub
PH_PY001_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY001_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY001_DataValidCheck() As Boolean
On Error GoTo PH_PY001_DataValidCheck_Error
    PH_PY001_DataValidCheck = False
    Dim i       As Long
    Dim j       As Long
    
    Dim sQry    As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    '//----------------------------------------------------------------------------------
    '//기본사항 탭
    '//----------------------------------------------------------------------------------
    If Trim(oDS_PH_PY001A.GetValue("Code", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "사원번호는 필수입니다.", bmt_Short, True
        oForm.Items("Code").CLICK ct_Regular
        PH_PY001_DataValidCheck = False
        Exit Function
    End If

    If oForm.Mode = fm_ADD_MODE Then
        '//사원번호 중복 체크 (PH_PY001A)
        sQry = "SELECT Code FROM [@PH_PY001A] WHERE Code = '" & Trim(oDS_PH_PY001A.GetValue("Code", 0)) & "'"
        oRecordSet.DoQuery sQry

        If oRecordSet.RecordCount > 0 Then
            Sbo_Application.SetStatusBarMessage "[PH_PY001A] 이미 존재하는 사원번호입니다", bmt_Short, True
            oForm.Items("Code").CLICK ct_Regular
            PH_PY001_DataValidCheck = False
            Exit Function
        End If

        '//사원번호 중복 체크(OHEM)
        sQry = "SELECT U_MSTCOD FROM OHEM WHERE U_MSTCOD = '" & Trim(oDS_PH_PY001A.GetValue("Code", 0)) & "'"
        oRecordSet.DoQuery sQry

        If oRecordSet.RecordCount > 0 Then
            Sbo_Application.SetStatusBarMessage "[OHEM] 이미 존재하는 사원번호입니다", bmt_Short, True
            oForm.Items("Code").CLICK ct_Regular
            PH_PY001_DataValidCheck = False
            Exit Function
        End If
        
        '// Code = Name 동일
    Call oDS_PH_PY001A.setValue("Name", 0, oDS_PH_PY001A.GetValue("Code", 0))
    End If
    

    If Trim(oDS_PH_PY001A.GetValue("U_CLTCOD", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "사업장은 필수입니다.", bmt_Short, True
        oForm.Items("CLTCOD").CLICK ct_Regular
        PH_PY001_DataValidCheck = False
        Exit Function
    End If

    If Trim(oDS_PH_PY001A.GetValue("U_StartDat", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "입사일은 필수입니다.", bmt_Short, True
        oForm.Items("startDat").CLICK ct_Regular
        PH_PY001_DataValidCheck = False
        Exit Function
    End If

    If Trim(oDS_PH_PY001A.GetValue("U_lNameKO", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "한글성 은 필수입니다.", bmt_Short, True
        oForm.Items("lNameKO").CLICK ct_Regular
        PH_PY001_DataValidCheck = False
        Exit Function
    End If
    
    If Trim(oDS_PH_PY001A.GetValue("U_fNameKO", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "한글명 은 필수입니다.", bmt_Short, True
        oForm.Items("fNameKO").CLICK ct_Regular
        PH_PY001_DataValidCheck = False
        Exit Function
    End If
    
    '//주민번호 체크
    If Trim(oDS_PH_PY001A.GetValue("U_govID", 0)) <> "" Then
        If GovIDCheck(Trim(oDS_PH_PY001A.GetValue("U_govID", 0))) = False Then
            Sbo_Application.SetStatusBarMessage "잘못된 주민번호입니다", bmt_Short, True
            oForm.Items("govID").CLICK ct_Regular
            PH_PY001_DataValidCheck = False
            Exit Function
        End If
    Else
        Sbo_Application.SetStatusBarMessage "주민번호 필수입니다.", bmt_Short, True
        oForm.Items("govID").CLICK ct_Regular
        PH_PY001_DataValidCheck = False
        Exit Function
    End If

    If Trim(oDS_PH_PY001A.GetValue("U_birthDat", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "생년월일은 필수입니다.", bmt_Short, True
        oForm.Items("birthDat").CLICK ct_Regular
        PH_PY001_DataValidCheck = False
        Exit Function
    End If

    If Trim(oDS_PH_PY001A.GetValue("U_fNameKO", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "한글명은 필수입니다.", bmt_Short, True
        oForm.Items("fNameKO").CLICK ct_Regular
        PH_PY001_DataValidCheck = False
        Exit Function
    End If
'    '//주민번호 중복 체크
'    sQry = "SELECT govID FROM OHEM WHERE govID = '" & Trim(oDS_PH_PY001A.GetValue("govID",0)) & "'"
'    oRecordset.DoQuery sQry
'
'    If oRecordset.RecordCount > 0 Then
'        Sbo_Application.SetStatusBarMessage "이미 존재하는 주민번호입니다", bmt_Short, True
'        oForm.Items("Code").CLICK ct_Regular
'        PH_PY001_DataValidCheck = False
'        Exit Function
'    End If

    If Trim(oDS_PH_PY001A.GetValue("U_sex", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "성별은 필수입니다.", bmt_Short, True
        oForm.Items("sex").CLICK ct_Regular
        PH_PY001_DataValidCheck = False
        Exit Function
    End If

    If Trim(oDS_PH_PY001A.GetValue("U_status", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "재직구분은 필수입니다.", bmt_Short, True
        If oForm.PaneLevel <> 1 Then oForm.PaneLevel = 1
        oForm.Items("status").CLICK ct_Regular
        PH_PY001_DataValidCheck = False
        Exit Function
    End If

    If Trim(oDS_PH_PY001A.GetValue("U_position", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "직책은 필수입니다.", bmt_Short, True
        If oForm.PaneLevel <> 1 Then oForm.PaneLevel = 1
        oForm.Items("position").CLICK ct_Regular
        PH_PY001_DataValidCheck = False
        Exit Function
    End If

    If Trim(oDS_PH_PY001A.GetValue("U_brthCntr", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "국적은 필수입니다.", bmt_Short, True
        If oForm.PaneLevel <> 1 Then oForm.PaneLevel = 1
        oForm.Items("brthCntr").CLICK ct_Regular
        PH_PY001_DataValidCheck = False
        Exit Function
    End If

    '//----------------------------------------------------------------------------------
    '//급여기본 탭
    '//----------------------------------------------------------------------------------
'    If oDS_PH_PY001A.GetValue("U_PAYTYP", 0) = "" Then
'        Sbo_Application.SetStatusBarMessage "급여형태는 필수입니다", bmt_Short, True
'        If oForm.PaneLevel <> 2 Then oForm.PaneLevel = 2
'        oForm.Items("PAYTYP").CLICK ct_Regular
'        PH_PY001_DataValidCheck = False
'        Exit Function
'    End If

    If oDS_PH_PY001A.GetValue("U_PAYSEL", 0) = "" Then
        Sbo_Application.SetStatusBarMessage "급여지급대상은 필수입니다", bmt_Short, True
        If oForm.PaneLevel <> 2 Then oForm.Items("FLD02").CLICK ct_Regular
        oForm.Items("PAYSEL").CLICK ct_Regular
        PH_PY001_DataValidCheck = False
        Exit Function
    End If

    If oDS_PH_PY001A.GetValue("U_STDAMT", 0) = "" Then
        Sbo_Application.SetStatusBarMessage "급여기본금은 필수입니다", bmt_Short, True
        If oForm.PaneLevel <> 2 Then oForm.PaneLevel = 2
        oForm.Items("STDAMT").CLICK ct_Regular
        PH_PY001_DataValidCheck = False
        Exit Function
    End If
    
    If oDS_PH_PY001A.GetValue("U_HUSMAN", 0) = "" Then
        Sbo_Application.SetStatusBarMessage "세대주여부는 필수입니다", bmt_Short, True
        If oForm.PaneLevel <> 2 Then oForm.PaneLevel = 2
        oForm.Items("HUSMAN").CLICK ct_Regular
        PH_PY001_DataValidCheck = False
        Exit Function
    End If
    '//----------------------------------------------------------------------------------
    '//급여기본 탭 - 고정수당
    '//----------------------------------------------------------------------------------
    If oMat1.VisualRowCount > 1 Then
        For i = 1 To oMat1.VisualRowCount - 1
            If oMat1.Columns("FILD01").Cells(i).Specific.VALUE <> "" Then
                If oMat1.Columns("FILD02").Cells(i).Specific.VALUE = "" Then
                    If oForm.PaneLevel <> 2 Then oForm.PaneLevel = 2
                    Sbo_Application.SetStatusBarMessage "수당명은 필수입니다", bmt_Short, True
                    PH_PY001_DataValidCheck = False
                    oMat1.Columns("FILD02").Cells(i).CLICK ct_Regular
                    Exit Function
                End If
'                If oMat1.Columns(i).Cells(3).Specific.Value = 0 Then
'                    If oForm.PaneLevel <> 2 Then oForm.PaneLevel = 2
'                    Sbo_Application.SetStatusBarMessage "수당금액 필수입니다", bmt_Short, True
'                    PH_PY001_DataValidCheck = False
'                    oMat1.Columns(i).Cells(3).CLICK ct_Regular
'                    Exit Function
'                End If
            End If
        Next i
    End If
    '//----------------------------------------------------------------------------------
    '//급여기본 탭 - 고정공제
    '//----------------------------------------------------------------------------------
    If oMat2.VisualRowCount > 1 Then
        For i = 1 To oMat2.VisualRowCount - 1
            If oMat2.Columns("FILD01").Cells(i).Specific.VALUE <> "" Then
                If oMat2.Columns("FILD02").Cells(i).Specific.VALUE = "" Then
                    If oForm.PaneLevel <> 2 Then oForm.PaneLevel = 2
                    Sbo_Application.SetStatusBarMessage "공제명은 필수입니다", bmt_Short, True
                    PH_PY001_DataValidCheck = False
                    oMat2.Columns("FILD02").Cells(2).CLICK ct_Regular
                    Exit Function
                End If
'                If oMat2.Columns(i).Cells(3).Specific.Value = 0 Then
'                    If oForm.PaneLevel <> 2 Then oForm.PaneLevel = 2
'                    Sbo_Application.SetStatusBarMessage "공제금액 필수입니다", bmt_Short, True
'                    PH_PY001_DataValidCheck = False
'                    oMat2.Columns(i).Cells(3).CLICK ct_Regular
'                    Exit Function
'                End If
            End If
        Next i
    End If
    
    '//----------------------------------------------------------------------------------
    '//기타사항 탭
    '//----------------------------------------------------------------------------------
    
    '//----------------------------------------------------------------------------------
    '//가족사항 탭
    '//----------------------------------------------------------------------------------
    '//가족 이름이 등록시...주민번호 입력 체크
    If oMat3.VisualRowCount > 0 Then
        For i = 1 To oMat3.VisualRowCount
            If oMat3.Columns("FamNam").Cells(i).Specific.VALUE <> "" Then
                If oMat3.Columns("FamPer").Cells(i).Specific.VALUE = "" Then
                    If oForm.PaneLevel <> 3 Then oForm.PaneLevel = 3
                    Sbo_Application.SetStatusBarMessage "입력된 가족의 주민등록번호는 필수입니다., bmt_Short, True"
                    PH_PY001_DataValidCheck = False
                    Exit Function
                End If
            End If
        Next i
    End If
    
    '//----------------------------------------------------------------------------------
    '//경력사항 탭
    '//----------------------------------------------------------------------------------
    If oMat4.VisualRowCount > 1 Then
        For i = 1 To oMat4.VisualRowCount - 1
            If oMat4.Columns("employer").Cells(i).Specific.VALUE <> "" Then
                If oMat4.Columns("fromDate").Cells(i).Specific.VALUE = "" Then
                    If oForm.PaneLevel <> 4 Then oForm.PaneLevel = 4
                    Sbo_Application.SetStatusBarMessage "기간시작은 필수입니다", bmt_Short, True
                    PH_PY001_DataValidCheck = False
                    oMat4.Columns(i).Cells(i).CLICK ct_Regular
                    Exit Function
                End If
                If oMat4.Columns("toDate").Cells(i).Specific.VALUE = "" Then
                    If oForm.PaneLevel <> 4 Then oForm.PaneLevel = 4
                    Sbo_Application.SetStatusBarMessage "기간종료는 필수입니다", bmt_Short, True
                    PH_PY001_DataValidCheck = False
                    oMat4.Columns("toDate").Cells(i).CLICK ct_Regular
                    Exit Function
                End If
            End If
        Next i
    End If
    
    '//----------------------------------------------------------------------------------
    '//자격사항탭
    '//----------------------------------------------------------------------------------
    
    '//----------------------------------------------------------------------------------
    '//발령사항 탭
    '//----------------------------------------------------------------------------------
    If oMat6.VisualRowCount > 1 Then
        For i = 1 To oMat6.VisualRowCount - 1
            If oMat6.Columns("appType").Cells(i).Specific.VALUE <> "" Then
                If oMat6.Columns("appDate").Cells(i).Specific.VALUE = "" Then
                    If oForm.PaneLevel <> 6 Then oForm.PaneLevel = 6
                    Sbo_Application.SetStatusBarMessage "발령일자는 필수입니다", bmt_Short, True
                    PH_PY001_DataValidCheck = False
                    oMat6.Columns("appDate").Cells(i).CLICK ct_Regular
                    Exit Function
                End If
                If oMat6.Columns("appNum").Cells(i).Specific.VALUE = "" Then
                    If oForm.PaneLevel <> 6 Then oForm.PaneLevel = 6
                    Sbo_Application.SetStatusBarMessage "발령번호는 필수입니다", bmt_Short, True
                    PH_PY001_DataValidCheck = False
                    oMat6.Columns(i).Cells(i).CLICK ct_Regular
                    Exit Function
                End If
            End If
        Next i
    End If
    
    '//----------------------------------------------------------------------------------
    '//동호회 탭
    '//----------------------------------------------------------------------------------
    
    '//----------------------------------------------------------------------------------
    '//항력사항 탭
    '//----------------------------------------------------------------------------------
    
    '//----------------------------------------------------------------------------------
    '//교육사항 탭
    '//----------------------------------------------------------------------------------
    If oMat9.VisualRowCount > 1 Then
        For i = 1 To oMat9.VisualRowCount - 1
            If oMat9.Columns("major").Cells(i).Specific.VALUE <> "" Then
                If oMat9.Columns("type").Cells(i).Specific.VALUE = "" Then
                    If oForm.PaneLevel <> 9 Then oForm.PaneLevel = 9
                    Sbo_Application.SetStatusBarMessage "교육사항 - 분류는 필수입니다", bmt_Short, True
                    PH_PY001_DataValidCheck = False
                    oMat9.Columns("type").Cells(i).CLICK ct_Regular
                    Exit Function
                End If
                If oMat9.Columns("fromDate").Cells(i).Specific.VALUE = "" Then
                    If oForm.PaneLevel <> 9 Then oForm.PaneLevel = 9
                    Sbo_Application.SetStatusBarMessage "교육사항 -  기간시작은 필수입니다", bmt_Short, True
                    PH_PY001_DataValidCheck = False
                    oMat9.Columns("fromDate").Cells(i).CLICK ct_Regular
                    Exit Function
                End If
                If oMat9.Columns("toDate").Cells(i).Specific.VALUE = 0 Then
                    If oForm.PaneLevel <> 9 Then oForm.PaneLevel = 9
                    Sbo_Application.SetStatusBarMessage "교육사항 - 기간종료는 필수입니다", bmt_Short, True
                    PH_PY001_DataValidCheck = False
                    oMat9.Columns("toDate").Cells(i).CLICK ct_Regular
                    Exit Function
                End If
            End If
        Next i
    End If
    '//----------------------------------------------------------------------------------
    '//상벌사항 탭
    '//----------------------------------------------------------------------------------
    
    '//----------------------------------------------------------------------------------
    '//여권사항 탭
    '//----------------------------------------------------------------------------------
    
    '//----------------------------------------------------------------------------------
    '//노조이력 탭
    '//----------------------------------------------------------------------------------
    
    '//----------------------------------------------------------------------------------
    '//경조사항 탭
    '//----------------------------------------------------------------------------------
    
    '//----------------------------------------------------------------------------------
    '//사고현황 탭
    '//----------------------------------------------------------------------------------
    
    '//DI API
    If oForm.Mode = fm_ADD_MODE Then
        '//OHEM 등록
        If PH_PY001_OHEM_DI_ADD = False Then
            PH_PY001_DataValidCheck = False
            Exit Function
        End If
    ElseIf oForm.Mode = fm_UPDATE_MODE Then
        If PH_PY001_OHEM_DI_UPDATE = False Then
            PH_PY001_DataValidCheck = False
            Exit Function
        End If
    End If
    
    oMat1.FlushToDataSource
    oMat2.FlushToDataSource
    oMat3.FlushToDataSource
    oMat4.FlushToDataSource
    oMat5.FlushToDataSource
    oMat6.FlushToDataSource
    oMat7.FlushToDataSource
    oMat8.FlushToDataSource
    oMat9.FlushToDataSource
    oMat10.FlushToDataSource
    oMat11.FlushToDataSource
    oMat12.FlushToDataSource
    oMat13.FlushToDataSource
    oMat14.FlushToDataSource
    oMat15.FlushToDataSource
    
    '// Matrix 마지막 행 삭제(DB 저장시)
    If oDS_PH_PY001B.Size > 1 Then oDS_PH_PY001B.RemoveRecord (oDS_PH_PY001B.Size - 1)
    If oDS_PH_PY001C.Size > 1 Then oDS_PH_PY001C.RemoveRecord (oDS_PH_PY001C.Size - 1)
    If oDS_PH_PY001D.Size > 1 Then oDS_PH_PY001D.RemoveRecord (oDS_PH_PY001D.Size - 1)
    If oDS_PH_PY001E.Size > 1 Then oDS_PH_PY001E.RemoveRecord (oDS_PH_PY001E.Size - 1)
    If oDS_PH_PY001F.Size > 1 Then oDS_PH_PY001F.RemoveRecord (oDS_PH_PY001F.Size - 1)
    If oDS_PH_PY001G.Size > 1 Then oDS_PH_PY001G.RemoveRecord (oDS_PH_PY001G.Size - 1)
    If oDS_PH_PY001H.Size > 1 Then oDS_PH_PY001H.RemoveRecord (oDS_PH_PY001H.Size - 1)
    If oDS_PH_PY001J.Size > 1 Then oDS_PH_PY001J.RemoveRecord (oDS_PH_PY001J.Size - 1)
    If oDS_PH_PY001K.Size > 1 Then oDS_PH_PY001K.RemoveRecord (oDS_PH_PY001K.Size - 1)
    If oDS_PH_PY001L.Size > 1 Then oDS_PH_PY001L.RemoveRecord (oDS_PH_PY001L.Size - 1)
    If oDS_PH_PY001M.Size > 1 Then oDS_PH_PY001M.RemoveRecord (oDS_PH_PY001M.Size - 1)
    If oDS_PH_PY001N.Size > 1 Then oDS_PH_PY001N.RemoveRecord (oDS_PH_PY001N.Size - 1)
    If oDS_PH_PY001P.Size > 1 Then oDS_PH_PY001P.RemoveRecord (oDS_PH_PY001P.Size - 1)
    If oDS_PH_PY001Q.Size > 1 Then oDS_PH_PY001Q.RemoveRecord (oDS_PH_PY001Q.Size - 1)
    If oDS_PH_PY001Q.Size > 1 Then oDS_PH_PY001R.RemoveRecord (oDS_PH_PY001R.Size - 1)
    
    oMat1.LoadFromDataSource
    oMat2.LoadFromDataSource
    oMat3.LoadFromDataSource
    oMat4.LoadFromDataSource
    oMat5.LoadFromDataSource
    oMat6.LoadFromDataSource
    oMat7.LoadFromDataSource
    oMat8.LoadFromDataSource
    oMat9.LoadFromDataSource
    oMat10.LoadFromDataSource
    oMat11.LoadFromDataSource
    oMat12.LoadFromDataSource
    oMat13.LoadFromDataSource
    oMat14.LoadFromDataSource
    oMat15.LoadFromDataSource
    PH_PY001_DataValidCheck = True
    
    Set oRecordSet = Nothing
    Exit Function
    
PH_PY001_DataValidCheck_Error:
    Set oRecordSet = Nothing
    PH_PY001_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY001_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY001_MTX01()

    '//메트릭스에 데이터 로드
    
    Dim i       As Long
    Dim sQry    As String
    
    Dim Param01 As String
    Dim Param02 As String
    Dim Param03 As String
    Dim Param04 As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY001_MTX01_Error

    Call oForm.Freeze(True)
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    Param01 = oForm.Items("Param01").Specific.VALUE
    Param02 = oForm.Items("Param01").Specific.VALUE
    Param03 = oForm.Items("Param01").Specific.VALUE
    Param04 = oForm.Items("Param01").Specific.VALUE
    
    sQry = "SELECT 10"
    Call oRecordSet.DoQuery(sQry)
    
    oMat1.Clear
    oMat1.FlushToDataSource
    oMat1.LoadFromDataSource
    
    If (oRecordSet.RecordCount = 0) Then
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
        GoTo PH_PY001_MTX01_Exit
    End If
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet.RecordCount, False)
    
    For i = 0 To oRecordSet.RecordCount - 1
        If i <> 0 Then
            oDS_PH_PY001B.InsertRecord (i)
        End If
        oDS_PH_PY001B.Offset = i
        Call oDS_PH_PY001B.setValue("U_COL01", i, oRecordSet.Fields(0).VALUE)
        Call oDS_PH_PY001B.setValue("U_COL02", i, oRecordSet.Fields(1).VALUE)
        oRecordSet.MoveNext
        ProgressBar01.VALUE = ProgressBar01.VALUE + 1
        ProgressBar01.Text = ProgressBar01.VALUE & "/" & oRecordSet.RecordCount & "건 조회중...!"
    Next
    oMat1.LoadFromDataSource
    oMat1.AutoResizeColumns
    oForm.Update
    
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Sub
PH_PY001_MTX01_Exit:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Exit Sub
PH_PY001_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY001_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY001_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY001_Validate_Error
    PH_PY001_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY001A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.VALUE & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", bmt_Short, True
        PH_PY001_Validate = False
        GoTo PH_PY001_Validate_Exit
    End If
'
    If ValidateType = "수정" Then

    ElseIf ValidateType = "행삭제" Then

    ElseIf ValidateType = "취소" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY001_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY001_Validate_Error:
    PH_PY001_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY001_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

'//인사기본&급여 저장 DI
Private Function PH_PY001_OHEM_DI_ADD() As Boolean
    Dim i                   As Long
    Dim IErrCode            As Long
    Dim sErrMsg             As String
    Dim sQry                As String
    
    Dim sItemName           As String
    
    Dim oOHEM               As SAPbobsCOM.EmployeesInfo
    Dim oRecordSet          As SAPbobsCOM.Recordset
    
    On Error GoTo Error_Handle
    
    PH_PY001_OHEM_DI_ADD = False
    
    Set oOHEM = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oEmployeesInfo)
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Sbo_Application.SetStatusBarMessage oForm.Items.Item("Code").Specific.VALUE & "사원마스터가 생성중입니다.", bmt_Short, False
    
    With oOHEM
        .UserFields.Fields.Item("U_MSTCOD").VALUE = oForm.Items.Item("Code").Specific.VALUE        '// 사원번호
        .UserFields.Fields.Item("U_CLTCOD").VALUE = oForm.Items.Item("CLTCOD").Specific.VALUE       '//사업장
        .UserFields.Fields.Item("U_GRPDAT").VALUE = Format(Trim(oForm.Items.Item("GrpDat").Specific.VALUE), "####-##-##")       '//그룹입사일
        .StartDate = Format(Trim(oForm.Items.Item("startDat").Specific.VALUE), "####-##-##")                                    '//입사일자
        If Format(Trim(oForm.Items.Item("RetDat").Specific.VALUE), "####-##-##") <> "" Then
            .UserFields.Fields.Item("U_RETDAT").VALUE = Format(Trim(oForm.Items.Item("RetDat").Specific.VALUE), "####-##-##") '//중간정산일 (값이 있을 때만 입력)
        End If
        .LastName = oForm.Items.Item("lNameKO").Specific.VALUE                                     '//한글 성
        .FirstName = oForm.Items.Item("fNameKO").Specific.VALUE                                    '//한글 명
        .UserFields.Fields.Item("U_FULLNAME").VALUE = Trim(.LastName) & Trim(.FirstName)           '//성명(풀네임)
        .UserFields.Fields.Item("U_lNameTW").VALUE = oForm.Items.Item("lNameTW").Specific.VALUE    '//한자 성
        .UserFields.Fields.Item("U_fNameTW").VALUE = oForm.Items.Item("fNameTW").Specific.VALUE    '//한자 명
        .UserFields.Fields.Item("U_lNameEN").VALUE = oForm.Items.Item("lNameEN").Specific.VALUE    '//영문 성
        .UserFields.Fields.Item("U_fNameEN").VALUE = oForm.Items.Item("fNameEN").Specific.VALUE    '//영문 명
        .IdNumber = oForm.Items.Item("govID").Specific.VALUE                                       '//주민번호(govID)
    
        If oForm.Items.Item("sex").Specific.VALUE = "M" Then                                        '//성별
            .Gender = gt_Male
        Else '// "F"
            .Gender = gt_Female
        End If
     
        .DateOfBirth = Format(Trim(oForm.Items.Item("birthDat").Specific.VALUE), "####-##-##")     '//생년월일
        .UserFields.Fields.Item("U_lastEduc").VALUE = oForm.Items.Item("lastEduc").Specific.VALUE  '//최종학력
        .eMail = oForm.Items.Item("eMail").Specific.VALUE                                          '//전자메일
        .HomeStreet = oForm.Items.Item("Address").Specific.VALUE                                   '//주소
        .HomeBlock = oForm.Items.Item("OriginAd").Specific.VALUE                                   '//본적
        .Remarks = oForm.Items.Item("MainJob").Specific.VALUE                                      '//주요업무
        .UserFields.Fields.Item("U_TeamCode").VALUE = oForm.Items.Item("TeamCode").Specific.VALUE  '//부서
        .UserFields.Fields.Item("U_RspCode").VALUE = oForm.Items.Item("RspCode").Specific.VALUE    '//담당
        .UserFields.Fields.Item("U_ClsCode").VALUE = oForm.Items.Item("ClsCode").Specific.VALUE    '//반
        .HomePhone = oForm.Items.Item("homeTel").Specific.VALUE                                    '//전화번호
        .OfficePhone = oForm.Items.Item("OffcTel").Specific.VALUE                                  '//사무실전화
        .MobilePhone = oForm.Items.Item("mobile").Specific.VALUE                                   '//휴대전화
        .StatusCode = oForm.Items.Item("status").Specific.VALUE                                    '//재직구분
    
        If oForm.Items.Item("status").Specific.VALUE = "5" Then                                     '//재직구분이 - 퇴직일 경우
            .TerminationDate = Format(Trim(oForm.Items.Item("termDate").Specific.VALUE), "####-##-##")         '//퇴직일자
        End If
    
        .Position = oForm.Items.Item("position").Specific.VALUE                                    '//직책
        .UserFields.Fields.Item("U_JIGCOD").VALUE = oForm.Items.Item("JIGCOD").Specific.VALUE      '//직급
        .UserFields.Fields.Item("U_JIGTYP").VALUE = oForm.Items.Item("JIGTYP").Specific.VALUE      '//직원구분
        .UserFields.Fields.Item("U_CallName").VALUE = oForm.Items.Item("CallName").Specific.VALUE  '//전문직호칭
        .UserFields.Fields.Item("U_preCode").VALUE = oForm.Items.Item("preCode").Specific.VALUE    '//변경전사번
    
        If oForm.Items.Item("martStat").Specific.VALUE = "Y" Then                                       '//혼인구분
            .MartialStatus = mts_Married
        Else '// "N"
            .MartialStatus = mts_Single
        End If
        
        If Trim(oForm.Items.Item("weddingD").Specific.VALUE) <> "" Then
            .UserFields.Fields.Item("U_weddingD").VALUE = Format(Trim(oForm.Items.Item("weddingD").Specific.VALUE), "####-##-##")  '//결혼기념일
        End If
        
        .CountryOfBirth = oForm.Items.Item("brthCntr").Specific.VALUE                              '//국적
        .UserFields.Fields.Item("U_ShiftTyp").VALUE = oForm.Items.Item("ShiftDat").Specific.VALUE  '//근무형태
        .UserFields.Fields.Item("U_GNMUJO").VALUE = oForm.Items.Item("GNMUJO").Specific.VALUE      '//근무조
        .UserFields.Fields.Item("U_KUKJUN").VALUE = oForm.Items.Item("KUKJUN").Specific.VALUE      '//퇴직근전환금
        
        
        '//경력사항 OHEM4
        If oMat4.VisualRowCount > 1 Then
            For i = 1 To oMat4.VisualRowCount - 1
                .PreviousEmpoymentInfo.Add
                .PreviousEmpoymentInfo.FromDtae = Format(oMat4.Columns("fromDate").Cells(i).Specific.VALUE, "####-##-##")           '//기간From
                .PreviousEmpoymentInfo.ToDate = Format(oMat4.Columns("toDate").Cells(i).Specific.VALUE, "####-##-##")               '//기간 To
                .PreviousEmpoymentInfo.Employer = oMat4.Columns("employer").Cells(i).Specific.VALUE                                 '//근무처
                .PreviousEmpoymentInfo.UserFields.Fields.Item("U_depart").VALUE = oMat4.Columns("depart").Cells(i).Specific.VALUE   '//부서
                .PreviousEmpoymentInfo.Position = oMat4.Columns("position").Cells(i).Specific.VALUE                                 '//직책
                .PreviousEmpoymentInfo.Remarks = oMat4.Columns("remarks").Cells(i).Specific.VALUE                                   '//직장소재지
            Next i
        End If
        
        '//교육사항 OHEM2
        If oMat9.VisualRowCount > 1 Then
            For i = 1 To oMat9.VisualRowCount - 1
                sQry = "SELECT Name FROM OHED WHERE edTYPE = '" & oMat9.Columns("type").Cells(i).Specific.VALUE & "'"
                oRecordSet.DoQuery sQry
    
                If oRecordSet.RecordCount > 0 Then
                    .EducationInfo.Add
                    .EducationInfo.EducationType = oMat9.Columns("type").Cells(i).Specific.VALUE                                    '//분류
                    .EducationInfo.Major = oMat9.Columns("major").Cells(i).Specific.VALUE                                           '//교육명
                    .EducationInfo.FromDate = Format(oMat9.Columns("fromDate").Cells(i).Specific.VALUE, "####-##-##")               '//기간from
                    .EducationInfo.ToDate = Format(oMat9.Columns("toDate").Cells(i).Specific.VALUE, "####-##-##")                   '//기간to
                    .EducationInfo.Institute = oMat9.Columns("institut").Cells(i).Specific.VALUE                                    '//교육기간
                    .EducationInfo.Diploma = oMat9.Columns("diploma").Cells(i).Specific.VALUE                                       '//교육장소
                    .EducationInfo.UserFields.Fields.Item("U_EduExp").VALUE = oMat9.Columns("EduExp").Cells(i).Specific.VALUE       '//교육비
                    .EducationInfo.UserFields.Fields.Item("U_TraExp").VALUE = oMat9.Columns("TraExp").Cells(i).Specific.VALUE       '//출장비
                Else
                    Sbo_Application.SetStatusBarMessage "등록된 교육 분류가 없어, 교육사항은 저장되지 않습니다", bmt_Short, False
                End If
            Next i
        End If
            
        If 0 <> oOHEM.Add Then
            oCompany.GetLastError IErrCode, sErrMsg
            Sbo_Application.SetStatusBarMessage sErrMsg, bmt_Short, True
            
        Else
            PH_PY001_OHEM_DI_ADD = True
            Sbo_Application.SetStatusBarMessage oForm.Items.Item("Code").Specific.VALUE & "사원마스터가 생성되었습니다.", bmt_Short, False
            
            Call oDS_PH_PY001A.setValue("U_FullName", 0, oOHEM.UserFields.Fields.Item("U_FULLNAME").VALUE)
        End If
        
    End With

    Set oOHEM = Nothing

    Exit Function

Error_Handle:

    Set oOHEM = Nothing
    Sbo_Application.SetStatusBarMessage "PH_PY001_OHEM_DI_ADD_err : " & Err.Number & " - " & Err.Description, bmt_Short, False
End Function


Private Function PH_PY001_OHEM_DI_UPDATE() As Boolean
    Dim i                   As Long
    Dim j                   As Long
    Dim IErrCode            As Long
    Dim sErrMsg             As String
    Dim sQry                As String
    Dim sItemName           As String
    Dim EmpID               As String
    
    Dim oRecordSet          As SAPbobsCOM.Recordset
    
    Dim oOHEM               As SAPbobsCOM.EmployeesInfo

    On Error GoTo Error_Handle

    PH_PY001_OHEM_DI_UPDATE = False
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    Set oOHEM = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oEmployeesInfo)
    
    sQry = "SELECT U_Empid FROM [@PH_PY001A] WHERE Code = '" & Trim(oForm.Items.Item("Code").Specific.VALUE) & "'"
    
    oRecordSet.DoQuery sQry
    
    EmpID = oRecordSet.Fields(0).VALUE

    
    Sbo_Application.SetStatusBarMessage oForm.Items.Item("Code").Specific.VALUE & "사원마스터가 갱신중입니다.", bmt_Short, False
    
    If EmpID = "" Then
        Sbo_Application.SetStatusBarMessage "empID 로드에 실패하였습니다.", bmt_Short, False
        GoTo Error_Handle
    End If
    If oOHEM.GetByKey(EmpID) = True Then
        
        With oOHEM
            .GetByKey (EmpID)
        
            .UserFields.Fields.Item("U_MSTCOD").VALUE = oForm.Items.Item("Code").Specific.VALUE        '// 사원번호
            .UserFields.Fields.Item("U_CLTCOD").VALUE = oForm.Items.Item("CLTCOD").Specific.VALUE       '//사업장
            .UserFields.Fields.Item("U_GRPDAT").VALUE = Format(Trim(oForm.Items.Item("GrpDat").Specific.VALUE), "####-##-##")       '//그룹입사일
            .StartDate = Format(Trim(oForm.Items.Item("startDat").Specific.VALUE), "####-##-##")                                 '//입사일자
            If Format(Trim(oForm.Items.Item("RetDat").Specific.VALUE), "####-##-##") <> "" Then
                .UserFields.Fields.Item("U_RETDAT").VALUE = Format(Trim(oForm.Items.Item("RetDat").Specific.VALUE), "####-##-##")      '//중간정산일
            End If
            .LastName = oForm.Items.Item("lNameKO").Specific.VALUE                                     '//한글 성
            .FirstName = oForm.Items.Item("fNameKO").Specific.VALUE                                    '//한글 명
            .UserFields.Fields.Item("U_FULLNAME").VALUE = Trim(.LastName) & Trim(.FirstName)           '//성명(풀네임)
            .UserFields.Fields.Item("U_lNameTW").VALUE = oForm.Items.Item("lNameTW").Specific.VALUE    '//한자 성
            .UserFields.Fields.Item("U_fNameTW").VALUE = oForm.Items.Item("fNameTW").Specific.VALUE    '//한자 명
            .UserFields.Fields.Item("U_lNameEN").VALUE = oForm.Items.Item("lNameEN").Specific.VALUE    '//영문 성
            .UserFields.Fields.Item("U_fNameEN").VALUE = oForm.Items.Item("fNameEN").Specific.VALUE    '//영문 명
            .IdNumber = oForm.Items.Item("govID").Specific.VALUE                                       '//주민번호(govID)
        
            If oForm.Items.Item("sex").Specific.VALUE = "M" Then                                        '//성별
                oOHEM.Gender = gt_Male
            Else '// "F"
                oOHEM.Gender = gt_Female
            End If
         
            .DateOfBirth = Format(Trim(oForm.Items.Item("birthDat").Specific.VALUE), "####-##-##")     '//생년월일
            .UserFields.Fields.Item("U_lastEduc").VALUE = oForm.Items.Item("lastEduc").Specific.VALUE  '//최종학력
            .eMail = oForm.Items.Item("eMail").Specific.VALUE                                          '//전자메일
            .HomeStreet = oForm.Items.Item("Address").Specific.VALUE                                   '//주소
            .HomeBlock = oForm.Items.Item("OriginAd").Specific.VALUE                                   '//본적
            .Remarks = oForm.Items.Item("MainJob").Specific.VALUE                                      '//주요업무
            .UserFields.Fields.Item("U_TeamCode").VALUE = oForm.Items.Item("TeamCode").Specific.VALUE  '//부서
            .UserFields.Fields.Item("U_RspCode").VALUE = oForm.Items.Item("RspCode").Specific.VALUE    '//담당
            .UserFields.Fields.Item("U_ClsCode").VALUE = oForm.Items.Item("ClsCode").Specific.VALUE    '//반
            .HomePhone = oForm.Items.Item("homeTel").Specific.VALUE                                    '//전화번호
            .MobilePhone = oForm.Items.Item("mobile").Specific.VALUE                                   '//휴대전화
            .StatusCode = oForm.Items.Item("status").Specific.VALUE                                    '//재직구분
        
            If oForm.Items.Item("status").Specific.VALUE = "5" Then  '//재직구분이 - 퇴직일 경우
                oOHEM.TerminationDate = Format(Trim(oForm.Items.Item("termDate").Specific.VALUE), "####-##-##")         '//퇴직일자
            End If
        
            .Position = oForm.Items.Item("position").Specific.VALUE                                    '//직책
            .UserFields.Fields.Item("U_JIGCOD").VALUE = oForm.Items.Item("JIGCOD").Specific.VALUE      '//직급
            .UserFields.Fields.Item("U_JIGTYP").VALUE = oForm.Items.Item("JIGTYP").Specific.VALUE      '//직원구분
            .UserFields.Fields.Item("U_CallName").VALUE = oForm.Items.Item("CallName").Specific.VALUE  '//전문직호칭
            .UserFields.Fields.Item("U_preCode").VALUE = oForm.Items.Item("preCode").Specific.VALUE    '//변경전사번
        
            If Trim(oForm.Items.Item("weddingD").Specific.VALUE) <> "" Then
                .UserFields.Fields.Item("U_weddingD").VALUE = Format(Trim(oForm.Items.Item("weddingD").Specific.VALUE), "####-##-##")  '//결혼기념일
            End If
        
            If oForm.Items.Item("martStat").Specific.VALUE = "Y" Then                                   '//혼인구분
                oOHEM.MartialStatus = mts_Married
            Else '// "N"
                oOHEM.MartialStatus = mts_Single
            End If
    
            .CountryOfBirth = oForm.Items.Item("brthCntr").Specific.VALUE                              '//국적
            .UserFields.Fields.Item("U_ShiftTyp").VALUE = oForm.Items.Item("ShiftDat").Specific.VALUE  '//근무형태
            .UserFields.Fields.Item("U_GNMUJO").VALUE = oForm.Items.Item("GNMUJO").Specific.VALUE      '//근무조
            .UserFields.Fields.Item("U_KUKJUN").VALUE = oForm.Items.Item("KUKJUN").Specific.VALUE      '//퇴직근전환금
            
            '경력사항 및 교육사항 DI는 불필요하여 주서처리함(2016.12.03 송명규)_S
'            '//경력사항 OHEM4
'            If oMat4.VisualRowCount > 1 Then
'                For i = 1 To oMat4.VisualRowCount - 1
'                    .PreviousEmpoymentInfo.Add
'                    .PreviousEmpoymentInfo.FromDtae = Format(oMat4.Columns("fromDate").Cells(i).Specific.VALUE, "####-##-##")           '//기간From
'                    .PreviousEmpoymentInfo.ToDate = Format(oMat4.Columns("toDate").Cells(i).Specific.VALUE, "####-##-##")               '//기간 To
'                    .PreviousEmpoymentInfo.Employer = oMat4.Columns("employer").Cells(i).Specific.VALUE                                 '//근무처
'                    .PreviousEmpoymentInfo.UserFields.Fields.Item("U_depart").VALUE = oMat4.Columns("depart").Cells(i).Specific.VALUE   '//부서
'                    .PreviousEmpoymentInfo.Position = oMat4.Columns("position").Cells(i).Specific.VALUE                                 '//직책
'                    .PreviousEmpoymentInfo.Remarks = oMat4.Columns("remarks").Cells(i).Specific.VALUE                                   '//직장소재지
'                Next i
'            End If
'
'            '//교육사항 OHEM2
'            If oMat9.VisualRowCount > 1 Then
'                For i = 1 To oMat9.VisualRowCount - 1
'                    sQry = "SELECT Name FROM OHED WHERE edTYPE = '" & oMat9.Columns("type").Cells(i).Specific.VALUE & "'"
'                    oRecordSet.DoQuery sQry
'
'                    If oRecordSet.RecordCount > 0 Then
'                        .EducationInfo.Add
'                        .EducationInfo.EducationType = oMat9.Columns("type").Cells(i).Specific.VALUE                                    '//분류
'                        .EducationInfo.Major = oMat9.Columns("major").Cells(i).Specific.VALUE                                           '//교육명
'                        .EducationInfo.FromDate = Format(oMat9.Columns("fromDate").Cells(i).Specific.VALUE, "####-##-##")               '//기간from
'                        .EducationInfo.ToDate = Format(oMat9.Columns("toDate").Cells(i).Specific.VALUE, "####-##-##")                   '//기간to
'                        .EducationInfo.Institute = oMat9.Columns("institut").Cells(i).Specific.VALUE                                    '//교육기간
'                        .EducationInfo.Diploma = oMat9.Columns("diploma").Cells(i).Specific.VALUE                                       '//교육장소
'                        .EducationInfo.UserFields.Fields.Item("U_EduExp").VALUE = oMat9.Columns("EduExp").Cells(i).Specific.VALUE       '//교육비
'                        .EducationInfo.UserFields.Fields.Item("U_TraExp").VALUE = oMat9.Columns("TraExp").Cells(i).Specific.VALUE       '//출장비
'                    Else
'                        Sbo_Application.SetStatusBarMessage "등록된 교육 분류가 없어, 교육사항은 저장되지 않습니다", bmt_Short, False
'                    End If
'                Next i
'            End If
            '경력사항 및 교육사항 DI는 불필요하여 주서처리함(2016.12.03 송명규)_E
        
            If 0 <> oOHEM.Update Then
                oCompany.GetLastError IErrCode, sErrMsg
                Sbo_Application.SetStatusBarMessage sErrMsg, bmt_Short, True
                
            Else
                PH_PY001_OHEM_DI_UPDATE = True
                Sbo_Application.SetStatusBarMessage oForm.Items.Item("Code").Specific.VALUE & "사원마스터가 갱신되었습니다.", bmt_Short, False
            End If
        End With
    End If
    
    Set oRecordSet = Nothing
    Set oOHEM = Nothing

    Exit Function

Error_Handle:
    Set oRecordSet = Nothing
    Set oOHEM = Nothing
    Sbo_Application.SetStatusBarMessage "PH_PY001_OHEM_DI_UPDATE_Error : " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Public Sub PH_PY001_empIDCheck()
    Dim sQry                As String
    Dim EmpID               As Long
    Dim MSTCOD              As String
    
    Dim oRecordSet          As SAPbobsCOM.Recordset
    
    On Error GoTo Error_Handle
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '/ OHEM - empID
    sQry = "SELECT Autokey FROM ONNM where objectcode = '171'"
    
    oRecordSet.DoQuery sQry
    
    EmpID = oRecordSet.Fields(0).VALUE - 1
    
    '/ OHEM - empID 의 사원번호
    sQry = "SELECT U_MSTCOD FROM [OHEM] WHeRE empID = '" & EmpID & "'"
    
    oRecordSet.DoQuery sQry
    
    MSTCOD = oRecordSet.Fields(0).VALUE
    
    '/ @PH_PY001A 에 empid 업데이트
    sQry = "UPDATE [@PH_PY001A] SET U_empID = '" & EmpID & "'"
    sQry = sQry & " WHERE Code = '" & MSTCOD & "'"
    
    oRecordSet.DoQuery sQry
    
    Set oRecordSet = Nothing
    
Error_Handle:
    Set oRecordSet = Nothing

End Sub


Private Sub PH_PY001_Display_Hobong()
On Error GoTo Error_Message
    Dim oRecordSet  As SAPbobsCOM.Recordset
    Dim sRecordset  As SAPbobsCOM.Recordset
    Dim sQry        As String
    Dim iRow        As Integer: Dim JSUCnt As Integer
    Dim jRow        As Integer
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    Set sRecordset = oCompany.GetBusinessObject(BoRecordset)
    
    sQry = " SELECT  U_CSUCOD, U_CSUNAM"
    sQry = sQry & " FROM    [@PH_PY102B] T1"
    If Trim$(oDS_PH_PY001A.GetValue("U_HOBYMM", 0)) <> "" Then
        sQry = sQry & " WHERE   CODE = (SELECT TOP 1 CODE FROM [@PH_PY102A] WHERE CODE <= '" & Trim$(oDS_PH_PY001A.GetValue("U_HOBYMM", 0)) & "' ORDER BY CODE DESC)"
    Else
        sQry = sQry & " WHERE   CODE = (SELECT TOP 1 CODE FROM [@PH_PY102A] ORDER BY CODE DESC)"
    End If
    sQry = sQry & " AND     U_HOBUSE = 'Y'"
   ' sQry = sQry & " AND     U_CSUGBN NOT IN ('10','11','18')"
    sQry = sQry & " ORDER   BY U_INSLIN"
    sRecordset.DoQuery sQry
    
    If sRecordset.RecordCount > 0 Then
    
        jRow = 0
        sQry = "SELECT TOP 1 T0.U_STDAMT, T0.U_BNSAMT, T0.U_EXTAMT01, T0.U_EXTAMT02, T0.U_EXTAMT03, T0.U_EXTAMT04, T0.U_EXTAMT05, "
        sQry = sQry & " T0.U_EXTAMT06, T0.U_EXTAMT07, T0.U_EXTAMT08, T0.U_EXTAMT09, T0.U_EXTAMT10, T1.U_YM"
        sQry = sQry & " FROM [@PH_PY105B] T0 INNER JOIN [@PH_PY105A] T1 ON T0.Code = T1.Code"
        sQry = sQry & " WHERE T0.U_JIGCOD = '" & Trim$(oDS_PH_PY001A.GetValue("U_JIGCOD", 0)) & "'"
        sQry = sQry & " AND   T0.U_HOBCOD = '" & Trim$(oDS_PH_PY001A.GetValue("U_HOBONG", 0)) & "'"
        If Trim$(oDS_PH_PY001A.GetValue("U_HOBYMM", 0)) <> "" Then
            sQry = sQry & " AND   T1.U_YM = '" & Trim$(oDS_PH_PY001A.GetValue("U_HOBYMM", 0)) & "'"
        End If
        sQry = sQry & " ORDER BY T1.U_YM DESC"
        
        oRecordSet.DoQuery sQry
        
        If oRecordSet.RecordCount > 0 Then
            oDS_PH_PY001A.setValue "U_STDAMT", 0, oRecordSet.Fields(0).VALUE
            oDS_PH_PY001A.setValue "U_BNSAMT", 0, oRecordSet.Fields(1).VALUE
            oDS_PH_PY001A.setValue "U_HOBYMM", 0, oRecordSet.Fields("U_YM").VALUE
            'JSUCnt = oRecordSet.Fields("U_JSUCNT").Value
            
            For iRow = 0 To oDS_PH_PY001B.Size - 1
'                oDS_PH_PY001B.Offset = iRow
'                If sRecordset.EOF Or jRow > JSUCnt Then
'                    Exit For
'                End If
                If Trim$(oDS_PH_PY001B.GetValue("U_FILD01", iRow)) = sRecordset.Fields("U_CSUCOD").VALUE Then
                    jRow = jRow + 1
                    oDS_PH_PY001B.setValue "U_FILD03", iRow, oRecordSet.Fields(1 + jRow).VALUE
                    sRecordset.MoveNext
                End If
            Next iRow
    '        For i = 1 To JSUCnt
    '            oDS_PH_PY001B.Offset = i - 1
    '            oDS_PH_PY001B.SetValue "U_FILD03", i - 1, oRecordSet.Fields(1 + i).Value '/ 금액
    '        Next i
    
        End If
        oForm.Items("STDAMT").Update
        oForm.Items("BNSAMT").Update
        oMat1.LoadFromDataSource
    End If
    
    Set oRecordSet = Nothing
    Set sRecordset = Nothing
    Exit Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Error_Message:
    Set oRecordSet = Nothing
    Set sRecordset = Nothing
    Sbo_Application.StatusBar.SetText "PH_PY001_Display_Hobong Error :" & Space$(10) & Err.Description, bmt_Short, smt_Error
End Sub

'*******************************************************************
'// oPaneLevel ==> 0:All / 1:oForm.PaneLevel=1 / 2:oForm.PaneLevel=2
'*******************************************************************
Private Sub PH_PY001_Matrix_Inti()
On Error GoTo Error_Message
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim i               As Long
    Dim cnt             As Integer
    Dim ErrNum          As Integer
    Dim DataChk         As Boolean
    Dim old_Line        As Integer
    ErrNum = 0
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '/ 수당Matrix 초기화
    cnt = oDS_PH_PY001B.Size
    If cnt > 0 Then
        For i = 1 To cnt - 1
            oDS_PH_PY001B.RemoveRecord oDS_PH_PY001B.Size - 1
        Next i
        If cnt = 1 Then
            oDS_PH_PY001B.Clear
        End If
    End If
    oMat1.LoadFromDataSource
    '/ 공제Matrix 초기화
    cnt = oDS_PH_PY001C.Size
    If cnt > 0 Then
        For i = 1 To cnt - 1
            oDS_PH_PY001C.RemoveRecord oDS_PH_PY001C.Size - 1
        Next i
        If cnt = 1 Then
            oDS_PH_PY001C.Clear
        End If
    End If
    oMat2.LoadFromDataSource
'// 기존 재직자의 수당항목이랑 비교해봅시당
    DataChk = True
    sQry = " SELECT  T0.U_FILD01, T0.U_FILD02"
    sQry = sQry & " FROM [@PH_PY001A] T1 INNER JOIN [@PH_PY001B] T0 ON T0.Code = T1.Code"
    sQry = sQry & " WHERE T1.Code = (SELECT TOP 1 Code"
    sQry = sQry & " FROM [@PH_PY001A] WHERE U_Status ='1'"
    sQry = sQry & " ORDER BY Code)"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        old_Line = 0
        Do Until oRecordSet.EOF
            If oRecordSet.Fields(0).VALUE <> "" Then
                old_Line = old_Line + 1
            End If
            oRecordSet.MoveNext
        Loop
        DataChk = False
    End If
'/ 수당항목코드읽어오기
    sQry = " SELECT T0.U_CSUCOD, T0.U_CSUNAM "
    sQry = sQry & " FROM [@PH_PY102B] T0 INNER JOIN [@PH_PY102A] T1 ON T0.Code = T1.Code"
    sQry = sQry & " WHERE T0.U_FIXGBN = '1'"    '/ 고정급
    sQry = sQry & " AND Left(T0.U_CSUCOD,1)  <> 'A'"    '/ 기본급제외
'    sQry = sQry & " AND ISNULL(T1.U_Status, '') = 'O'"  '/ 현적용수당항목
    sQry = sQry & " AND ISNULL(T0.U_INSLIN, '') <> ''"  '/ 인사마스터 순서
    sQry = sQry & " ORDER BY T0.Code, CAST(T0.U_INSLIN AS INT), CAST(T0.U_LINSEQ AS INT)"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        If DataChk = False And old_Line < oRecordSet.RecordCount Then
            ErrNum = 1
            GoTo Error_Message
        End If
        i = 1
        Do Until oRecordSet.EOF
            If i > oDS_PH_PY001B.Size Then
                oDS_PH_PY001B.InsertRecord (i - 1)
            End If
            oDS_PH_PY001B.Offset = i - 1
            oDS_PH_PY001B.setValue "U_LineNum", i - 1, i
            oDS_PH_PY001B.setValue "U_FILD01", i - 1, oRecordSet.Fields(0).VALUE '/ 코드
            oDS_PH_PY001B.setValue "U_FILD02", i - 1, oRecordSet.Fields(1).VALUE '/ 명
            oDS_PH_PY001B.setValue "U_FILD03", i - 1, "" '/ 금액
            i = i + 1
            oRecordSet.MoveNext
        Loop
    End If
    oMat1.LoadFromDataSource
    
'///공제항목
    DataChk = True
    sQry = " SELECT  T0.U_FILD01, T0.U_FILD02"
    sQry = sQry & " FROM [@PH_PY001A] T1 INNER JOIN [@PH_PY001C] T0 ON T0.Code = T1.Code"
    sQry = sQry & " WHERE T1.Code = (SELECT TOP 1 Code"
    sQry = sQry & " FROM [@PH_PY001A] "
    sQry = sQry & " WHERE U_Status ='1'"
    sQry = sQry & " ORDER BY Code)"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        old_Line = 0
        Do Until oRecordSet.EOF
            If oRecordSet.Fields(0).VALUE <> "" Then
                old_Line = old_Line + 1
            End If
            oRecordSet.MoveNext
        Loop
        DataChk = False
    End If

'/ 공제항목코드읽어오기
    sQry = " SELECT T0.U_CSUCOD, T0.U_CSUNAM "
    sQry = sQry & " FROM [@PH_PY103B] T0 INNER JOIN [@PH_PY103A] T1 ON T0.Code = T1.Code"
    sQry = sQry & " WHERE T0.U_FIXGBN = '1'"    '/ 고정공제
    sQry = sQry & " AND ISNULL(T0.U_INSLIN, '') <> ''"  '/ 인사마스터순서
    sQry = sQry & " AND Left(T0.U_CSUCOD,1)  <> 'A'"    '/ 기본급제외
'    sQry = sQry & " AND ISNULL(T1.U_Status, '') = 'O'"  '/ 현적용공제항목
    sQry = sQry & " ORDER BY T0.Code, CAST(T0.U_INSLIN AS INT)"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        If DataChk = False And old_Line < oRecordSet.RecordCount Then
            ErrNum = 2
            GoTo Error_Message
        End If
                
        i = 1
        Do Until oRecordSet.EOF
            If i > oDS_PH_PY001C.Size Then
                oDS_PH_PY001C.InsertRecord (i - 1)
            End If
            oDS_PH_PY001C.Offset = i - 1
            oDS_PH_PY001C.setValue "U_LineNum", i - 1, i
            oDS_PH_PY001C.setValue "U_FILD01", i - 1, oRecordSet.Fields(0).VALUE '/ 코드
            oDS_PH_PY001C.setValue "U_FILD02", i - 1, oRecordSet.Fields(1).VALUE '/ 명
            oDS_PH_PY001C.setValue "U_FILD03", i - 1, "" '/ 금액
            i = i + 1
            oRecordSet.MoveNext
        Loop
    End If
    oMat2.LoadFromDataSource
        
   Set oRecordSet = Nothing
    Exit Sub
'/////////////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
   Set oRecordSet = Nothing
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "PH_PY001_Matrix_Inti Error: 기존자료와 수당항목관리의 고정수당이 다릅니다. 고정수당/공제 생성으로 재직자의 고정수당을 변경 후 사용하십시오." & Err.Description, bmt_Short, smt_Error
    ElseIf ErrNum = 2 Then
        Sbo_Application.StatusBar.SetText "PH_PY001_Matrix_Inti Error: 기존자료와 공제항목관리의 고정공제이 다릅니다. 고정수당/공제 생성으로 재직자의 고정공제를 변경 후 사용하십시오." & Err.Description, bmt_Short, smt_Error
    Else
        Sbo_Application.StatusBar.SetText "PH_PY001_Matrix_Inti Error:" & Err.Description, bmt_Short, smt_Error
    End If
End Sub


Public Sub PH_PY001_LoadPic()
    Dim i               As Integer
    Dim sFilePath       As String
    Dim sFileName       As String

    Dim SaveFolders     As String
    Dim objFile         As Object
    Dim FSO             As FileSystemObject

    On Error GoTo PH_PY001_LoadPic_Error
    
    Set FSO = New FileSystemObject
    
    '// 사번입력 확인
    If Trim(oDS_PH_PY001A.GetValue("Code", 0)) = "" Then Exit Sub
    
    '//사진 폴더 생성
    'SaveFolders = "\\191.1.1.220\b1_shr\PathINI\PicturePS"
    SaveFolders = "\\191.1.1.220\HR_Pic"
'    If Dir(SaveFolders, vbDirectory) = "" Then MkDir SaveFolders '폴더가 없을 경우 폴더 만들기(2018.03.28 송명규 주석처리) 폴더는 수동으로 생성
    
    '//사진 불러오기
    sFilePath = FileListBoxForm.OpenDialog(FileListBoxForm, "*.BMP", "파일선택", "C:\")
    If sFilePath = "" Then Exit Sub
    sFilePath = Replace(sFilePath, Chr$(0), "")
    sFileName = Trim$(Mid(sFilePath, InStr(sFilePath, "\") + 1, Len(sFilePath) - 1))
    If UCase(Right(sFileName, 3)) <> "BMP" Then
        Sbo_Application.SetStatusBarMessage "BMP 확장자만 가능합니다.", bmt_Short, True
        Exit Sub
    End If
    
    
    '//서버에 기존 파일 체크
    For Each objFile In FSO.GetFolder(SaveFolders).Files
        If InStr(Trim$(UCase(objFile.Name)), UCase$(Trim(oDS_PH_PY001A.GetValue("Code", 0))) & ".BMP") > 0 Then
            If Sbo_Application.MessageBox("현재 사원의 사진 파일이 있습니다, 교체하시겠습니까?", 2, "Yes", "No") = 1 Then
                '//기존 파일 삭제
                Kill SaveFolders & "\" & objFile.Name
            Else
                Exit Sub
            End If
        End If
    Next objFile
    
    '//파일이름을 사번으로 변경후 파일 복사
    'Call FileCopy(sFileName, SaveFolders & "\" & Trim(oDS_PH_PY001A.GetValue("Code", 0)) & ".BMP")
    Call FileCopy(sFilePath, SaveFolders & "\" & Trim(oDS_PH_PY001A.GetValue("Code", 0)) & ".BMP")
    
    Sbo_Application.SetStatusBarMessage "사진이 서버에 업로드 되었습니다.", bmt_Short, False
    If oForm.Mode = fm_OK_MODE Then
        oForm.Mode = fm_UPDATE_MODE
    End If
    
    oDS_PH_PY001A.setValue "U_Pic1", 0, SaveFolders & "\" & Trim(oDS_PH_PY001A.GetValue("Code", 0)) & ".BMP"
    
    
    Set FSO = Nothing
    
    Exit Sub
    
PH_PY001_LoadPic_Error:
    Set FSO = Nothing
    Sbo_Application.SetStatusBarMessage "PH_PY001_LoadPic_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


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

