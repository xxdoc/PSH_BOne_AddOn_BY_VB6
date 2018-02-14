VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ZPY421"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : ZPY421.cls
'//  Module         : 인사관리 > 퇴직소득전산매체수록
'//  Desc           : 퇴직소득전산매체수록
'//********************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form

''// 그리드 사용시
'Public oGrid1           As SAPbouiCOM.Grid
'Public oDS_ZPY421     As SAPbouiCOM.DataTable
'
''// 매트릭스 사용시
Public oMat1 As SAPbouiCOM.Matrix

'Private oDS_ZPY421A As SAPbouiCOM.DBDataSource
'Private oDS_ZPY421B As SAPbouiCOM.DBDataSource

Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long


Private oJsnYear    As String
Private DPTSTR      As String
Private DPTEND      As String
Private MSTCOD      As String
Private CLTCOD      As String
Private BUSCNT      As Integer  '/ B레코드일련번호
Private BUSTOT      As Integer  '/ B레코드총갯수

Private oFilePath   As String

Private FILNAM      As String * 30    '파  일  명
Private MaxRow      As Long
Private NEWCNT      As Integer
Private C_MSTCOD    As String

'/ 퇴직소득 지급명세서-1
Private Type A_Record
        RECGBN As String * 1     '레코드구분
        DTAGBN As String * 2     '자료  구분
        TAXCOD As String * 3     '세  무  서
        PRTDAT As String * 8     '제출  일자
        
        RPTGBN As String * 1     '제  출  자 (1;세무대리인, 2;법인, 3;개인)
        TAXAGE As String * 6     '세무대리인
        HOMTID As String * 20    '홈텍스ID
        PGMCOD As String * 4     '세무프로그램코드
        BUSNBR As String * 10    '사업자번호
        SANGHO As String * 40    '상      호
        DAMDPT As String * 30    '담당부서명
        DAMNAM As String * 30    '담당 성명
        DAMTEL As String * 15    '담당전화번호
        
        BUSCNT As String * 5     'B Record수(신고의무자수)
        HANCOD As String * 3     '한글코드종
        FILLER As String * 812   '공      란
End Type
Dim Arec As A_Record


Private Type B_Record
        RECGBN As String * 1     '레코드구분
        DTAGBN As String * 2     '자료  구분
        TAXCOD As String * 3     '세  무  서
        BUSCNT As String * 6     '일련  번호
        
        BUSNBR As String * 10    '사업자번호
        SANGHO As String * 40    '상      호
        COMPRT As String * 30    '대  표  자
        PERNBR As String * 13    '주민  번호
        
        NEWCNT As String * 7     'C Record수
        OLDCNT As String * 7     'D Record수
        INCOME As String * 14    '소득금액총액
        GULGAB As String * 13    '결정갑근세
        GULCOM As String * 13    '    법인세
        GULJUM As String * 13    '    주민세
        GULNON As String * 13    '    농특세
        GULTOT As String * 13    '    총  액
        RNGCOD As String * 1     '제출대상기간
        FILLER As String * 791   '공      란
End Type
Dim Brec As B_Record

'/ 퇴직 주(현) 근무처 레코드 /
Private Type C_Record
        RECGBN As String * 1     '레코드구분
        DTAGBN As String * 2     '자료  구분
        TAXCOD As String * 3     '세  무  서
        SQNNBR As String * 6     '일련  번호
        BUSNBR As String * 10    '사업자번호
        
        BUSNUM As String * 10    '사업자번호(주현)
        SANGHO As String * 40    '상      호
        
        JONCNT As String * 2     '종전근무처수
        DWEGBN As String * 1     '거주자구분
        RGNCOD As String * 2     '거주지국코드
        STRINT As String * 8     '귀속년도시작
        ENDINT As String * 8     '귀속년도종료
        MSTNAM As String * 30    '성      명
        INTGBN As String * 1     '내외국인구분
        PERNBR As String * 13    '주민  번호
        
        RETRES As String * 1     '퇴직  사유
        TJKPAY As String * 10    '퇴직 급여(법정)
        SUDAMT As String * 10    '퇴직 급여(법정이외)-명예퇴직수당(추가퇴직금)
        BHMAMT As String * 10    '퇴직연금일시금(법정)(단체퇴직보험금)
        BHMAM1 As String * 10    '퇴직연금일시금(법정 외)           (2010년 추가)
        TOTAMT As String * 10    '퇴직급여총액(법정)
        TOTAM1 As String * 10    '퇴직급여총액(법정외 2011년추가)
        BTXPAY As String * 10    '비과세소득(2007년)
        
        MYNACC As String * 20    '퇴직연금 계좌번호                 (2010년 추가)
        MYNTOT As String * 10    '연금-총수령액(2007년)
        MYNWON As String * 10    '연금-월리금합계액(2007년)
        MYNBUL As String * 10    '연금-소득자불입액(2007년)
        MYNGON As String * 10    '연금-퇴직연금소득공제액(2007년)
        MYNYIL As String * 10    '연금-퇴직연금일시금(2007년)
        
        JYNACC As String * 20    '종(전) 퇴직연금 계좌번호          (2010년 추가)
        JYNTOT As String * 10    '종(전) 퇴직연금-총수령액          (2007년 추가/2010년 위치이동)
        JYNWON As String * 10    '종(전) 퇴직연금-원리금합계액      (2007년 추가/2010년 위치이동)
        JYNBUL As String * 10    '종(전) 퇴직연금-소득자불입액      (2007년 추가/2010년 위치이동)
        JYNGON As String * 10    '종(전) 퇴직연금-퇴직연금소득공제액(2007년 추가/2010년 위치이동)
        JYITOT As String * 10    '종(전) 퇴직연금-퇴직연금일시금    (2007년 추가/2010년 위치이동)
        
        SH1JIG As String * 10    '환산-지급예상액-주현              (2007년)
        SH3JIG As String * 10    '환산-지급예상액-종전              (2009년)
        SH1IYO As String * 10    '환산-과세이연금액-주현            (2010년 추가)
        SH3IYO As String * 10    '환산-과세이연금액-종전            (2010년 추가)
        SH1GIS As String * 10    '환산-기수령금액-주현              (2010년 추가)
        SH3GIS As String * 10    '환산-기수령금액-종전              (2010년 추가)
        SH1TIL As String * 10    '환산-총일시금                     (2007년)
        SH1SUR As String * 10    '환산-수령가능퇴직급여액           (2007년)
        SH1GON As String * 10    '환산-환산퇴직소득공제             (2007년)
        SH1GWA As String * 10    '환산-환산퇴직소득과세표준         (2007년)
        SH1YAG As String * 10    '환산-환산연평균과세표준           (2007년)
        SH1YAS As String * 10    '환산-연평균산출세액               (2007년)
        
        SH2JIG As String * 10    '환산-법정외-지급예상액-주현       (2007년)
        SH4JIG As String * 10    '환산-법정외-지급예상액-종전       (2009년)
        SH2IYO As String * 10    '환산-법정외-과세이연금액-주현     (2010년 추가)
        SH4IYO As String * 10    '환산-법정외-과세이연금액-종전     (2010년 추가)
        SH2GIS As String * 10    '환산-법정외-기수령금액-주현       (2010년 추가)
        SH4GIS As String * 10    '환산-법정외-기수령금액-종전       (2010년 추가)
        SH2TIL As String * 10    '환산-법정외-총일시금              (2007년)
        SH2SUR As String * 10    '환산-법정외-수령가능퇴직급여액    (2007년)
        SH2GON As String * 10    '환산-법정외-환산퇴직소득공제      (2007년)
        SH2GWA As String * 10    '환산-법정외-환산퇴직소득과세표준  (2007년)
        SH2YAG As String * 10    '환산-법정외-환산연평균과세표준    (2007년)
        SH2YAS As String * 10    '환산-법정외-연평균산출세액        (2007년)
        
        INPDAT As String * 8     '법정-현 입사일자
        OUTDAT As String * 8     '법정-현 퇴사일자
        GNMMON As String * 4     '법정-현 근속월수
        EXPMON As String * 4     '법정-현 제외월수(2007년)
        JIPDAT As String * 8     '법정-전 입사일자                  (2010년 위치이동)
        JOTDAT As String * 8     '법정-전 퇴사일자                  (2010년 위치이동)
        JGMMON As String * 4     '법정-전 근속월수                  (2010년 위치이동)
        JEXMON As String * 4     '법정-전 제외월수(2007년)          (2010년 위치이동)
        DUPMON As String * 4     '법정-   중복월수
        GNMYER As String * 2     '법정-   근속연수
        
        FR2DAT As String * 8     '법정외-현  입사일자
        TO2DAT As String * 8     '법정외-현  퇴사일자
        GN2MON As String * 4     '법정외-현  근속월수
        EX2MON As String * 4     '법정외-현  제외월수(2007년)
        JI2DAT As String * 8     '법정외-전  입사일자               (2010년 위치이동)
        JO2DAT As String * 8     '법정외-전  퇴사일자               (2010년 위치이동)
        JO2MON As String * 4     '법정외-전  근속월수               (2010년 위치이동)
        JE2MON As String * 4     '법정외-전  제외월수(2007년)       (2010년 위치이동)
        JB2MON As String * 4     '법정외-전  중복월수
        GN2YER As String * 2     '법정외-    근속연수
        
        JS1RET As String * 10    '법정-퇴직급여액
        JS1GON As String * 10    '법정-퇴직소득공제
        JS1STD As String * 10    '법정-퇴직소득과세표준
        JS1YSD As String * 10    '법정-연평균과세표준
        JS1YST As String * 10    '법정-연평균산출세액
        JS1SAN As String * 10    '법정-산출  세액
        JS1FRN As String * 10    '법정-세액(외국납부)공제 (세액공제2007년삭제=>2008년외국납부세액공제대체)
        
        JS2RET As String * 10    '법정외-퇴직급여액
        JS2GON As String * 10    '법정외-퇴직소득공제
        JS2STD As String * 10    '법정외-퇴직소득과세표준
        JS2YSD As String * 10    '법정외-연평균과세표준
        JS2YST As String * 10    '법정외-연평균산출세액
        JS2SAN As String * 10    '법정외-산출  세액
        JS2FRN As String * 10    '세액(외국납부)공제 (세액공제2007년삭제=>2008년외국납부세액공제대체)
        
        RETPAY As String * 10    '계-퇴직급여액
        RETGON As String * 10    '계-퇴직소득공제
        TAXSTD As String * 10    '계-퇴직소득과세표준
        YTXSTD As String * 10    '계-연평균과세표준
        YSANTX As String * 10    '계-연평균산출세액
        SANTAX As String * 10    '계-산출  세액
        TAXGON As String * 10     '계-세액(외국납부)공제 (세액공제2007년삭제=>2008년외국납부세액공제대체)
        
        GULGAB As String * 10    '결정소득세
        GULJUM As String * 10    '    주민세
        GULNON As String * 10    '    특별세
        GULTOT As String * 10    '    세액계
        
        JONGAB As String * 10    '종전소득세
        JONJUM As String * 10    '    주민세
        JONNON As String * 10    '    특별세
        JONTOT As String * 10    '    세액계
        FILLER As String * 4     '공      란
End Type
Dim Crec As C_Record

Private Type D_Record
        RECGBN As String * 1    '레코드구분
        DTAGBN As String * 2    '자료  구분
        TAXCOD As String * 3    '세  무  서
        SQNNBR As String * 6    '일련  번호
        
        BUSNBR As String * 10   '사업자번호
        FILLD1 As String * 50   '공      란
        
        PERNBR As String * 13   '주민  번호
        
        JONNAM As String * 40   '근무처  명
        JONNBR As String * 10   '사업자번호
        RETAMT As String * 10   '퇴직  급여
        SUDAMT As String * 10   '명예퇴직수당
        BHMAMT As String * 10   '퇴직연금일시금(법정)(퇴직보험금등)
        BHMAM1 As String * 10   '퇴직연금일시금(법정외)
        TOTAMT As String * 10   '종전계
        TOTAM1 As String * 10   '종전계(법정이외)
        BTXP01 As String * 10   '비과세소득(2007년 추가)
                
        JONCNT As String * 2    '종전근무처일련번호
        FILLER As String * 783  '공      란
End Type
Dim Drec As D_Record

Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\ZPY421.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "ZPY421_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "ZPY421"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
'    oForm.DataBrowser.BrowseBy = "Code"
        
    oForm.Freeze True
    Call ZPY421_CreateItems
    Call ZPY421_EnableMenus
    Call ZPY421_SetDocument(oFromDocEntry01)
'    Call ZPY421_FormResize
    
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

Private Function ZPY421_CreateItems() As Boolean

    Dim sQry        As String
    Dim i           As Long
    
    Dim oCheck      As SAPbouiCOM.CheckBox
    Dim oEdit       As SAPbouiCOM.EditText
    Dim oCombo1      As SAPbouiCOM.ComboBox
    Dim oCombo2      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns    As SAPbouiCOM.Columns
    Dim optBtn      As SAPbouiCOM.OptionBtn
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo ZPY421_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '/ 생성년도
    Call oForm.DataSources.UserDataSources.Add("STRDAT", dt_DATE, 10)
    Set oEdit = oForm.Items("STRDAT").Specific
    oEdit.DataBind.SetBound True, "", "STRDAT"
    
    '/ 생성년도
    Call oForm.DataSources.UserDataSources.Add("ENDDAT", dt_DATE, 10)
    Set oEdit = oForm.Items("ENDDAT").Specific
    oEdit.DataBind.SetBound True, "", "ENDDAT"
    
    'Call oForm.DataSources.UserDataSources.Add("JsnGbn", dt_SHORT_TEXT, 10)     '/ 생성구분
    
    '/ 생성구분
    Call oForm.DataSources.UserDataSources.Add("JSNTYP", dt_SHORT_TEXT, 10)
    Set oCombo1 = oForm.Items("JSNTYP").Specific
    oCombo1.ValidValues.Add "1", "연간(01.01~12.31)지급분"
    oCombo1.ValidValues.Add "2", "폐업에 의한 수시 제출분"
    oCombo1.ValidValues.Add "3", "수시 분할제출분"
    oCombo1.Select 0, psk_Index
    
    
    
    '/부서 From
    Call oForm.DataSources.UserDataSources.Add("DptStr", dt_SHORT_TEXT, 10)     '/ 부서코드
    Set oCombo1 = oForm.Items("DptStr").Specific
    
    '/부서 To
    Call oForm.DataSources.UserDataSources.Add("DptEnd", dt_SHORT_TEXT, 10)
    Set oCombo2 = oForm.Items("DptEnd").Specific
    
    '/사번
    Call oForm.DataSources.UserDataSources.Add("MSTCOD", dt_SHORT_TEXT, 8)
    Set oEdit = oForm.Items("MSTCOD").Specific
    oEdit.DataBind.SetBound True, "", "MSTCOD"
    
    '/성명
    Call oForm.DataSources.UserDataSources.Add("MSTNAM", dt_SHORT_TEXT, 30)
    Set oEdit = oForm.Items("MSTNAM").Specific
    oEdit.DataBind.SetBound True, "", "MSTNAM"
    
    '/사번순번
    Call oForm.DataSources.UserDataSources.Add("EmpID", dt_SHORT_TEXT, 10)
    Set oEdit = oForm.Items("EmpID").Specific
    oEdit.DataBind.SetBound True, "", "EmpID"
    
    '/제출일자
    Call oForm.DataSources.UserDataSources.Add("PRTDAT", dt_DATE, 10)
    Set oEdit = oForm.Items("PRTDAT").Specific
    oEdit.DataBind.SetBound True, "", "PRTDAT"
    
    '/경로
    Call oForm.DataSources.UserDataSources.Add("Path", dt_SHORT_TEXT)
    Set oEdit = oForm.Items("Path").Specific
    oEdit.DataBind.SetBound True, "", "Path"
    
    '// 생성구분
'    Set oCombo1 = oForm.Items("JsnGbn").Specific
'    oCombo1.ValidValues.Add "1", "퇴직소득 지급명세서-1"
'    oCombo1.ValidValues.Add "2", "퇴직소득 지급명세서-2"
'    Call oCombo1.Select(0, psk_Index)   '/ 전체
    
    '// 사업장
    Call oForm.DataSources.UserDataSources.Add("CLTCOD", dt_SHORT_TEXT, 10)      '/ 지점
    Set oCombo1 = oForm.Items("CLTCOD").Specific
    oCombo1.DataBind.SetBound True, "", "CLTCOD"
    
    Set oMat1 = oForm.Items("Mat1").Specific
    
    Call oForm.DataSources.UserDataSources.Add("Col0", dt_SHORT_TEXT)
    Call oForm.DataSources.UserDataSources.Add("Col1", dt_SHORT_TEXT)
    
    Set oColumn = oMat1.Columns("Col0")
    oColumn.DataBind.SetBound True, "", "Col0"
    
    Set oColumn = oMat1.Columns("Col1")
    oColumn.DataBind.SetBound True, "", "Col1"
    
    '//----------------------------------------------------------------------------------------------
    '// 데이터셋정의
    '//----------------------------------------------------------------------------------------------
'    '//테이블이 있을경우 데이터셋(Matrix)
'    Set oDS_ZPY421A = oForm.DataSources.DBDataSources("@ZPY421A")   '//헤더
'    Set oDS_ZPY421B = oForm.DataSources.DBDataSources("@ZPY421B")   '//라인
'
'    Set oMat1 = oForm.Items("Mat1").Specific       '
'
'    oMat1.SelectionMode = ms_NotSupported
'    oMat1.AutoResizeColumns

'    '//테이블이 없는경우 데이터셋(Grid)
'    oForm.DataSources.DataTables.Add ("PH_PY004")
'    oForm.DataSources.DataTables.Item("PH_PY004").Columns.Add "부서", ft_AlphaNumeric
'    oForm.DataSources.DataTables.Item("PH_PY004").Columns.Add "담당", ft_AlphaNumeric
'
'    Set oGrid1 = oForm.Items("Grid1").Specific
'
'    oGrid1.DataTable = oForm.DataSources.DataTables.Item("PH_PY004")
'    Set oDS_PH_PY004 = oForm.DataSources.DataTables.Item("PH_PY004")
    
    
    '//----------------------------------------------------------------------------------------------
    '// 아이템 설정
    '//----------------------------------------------------------------------------------------------
'    '//콤보1
'    Set oCombo1 = oForm.Items("    ").Specific
'    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
'    Call SetReDataCombo(oForm, sQry, oCombo1)
'    oForm.Items("    ").DisplayDesc = True
'
'    '//콤보2
'    Set oCombo1 = oForm.Items("    ").Specific
'    oCombo1.ValidValues.Add "M", "남자"
'    oCombo1.ValidValues.Add "F", "여자"
''    oCombo1.Select 0, psk_Index
'    oForm.Items("sex").DisplayDesc = True
'
'    '/체크박스
'    Set oCheck = oForm.Items("    ").Specific
'    oCheck.ValOn = "Y": oCheck.ValOff = "N"
'    oCheck.Checked = False
'
'    '//매트릭스컬럼
'    Set oColumn = oMat1.Columns("FILD01")
'    oColumn.Editable = True

'    '//UserDataSources
'    Call oForm.DataSources.UserDataSources.Add("     ", dt_SHORT_TEXT, 10)
'    Set oCombo1 = oForm.Items("    ").Specific
'    oCombo1.DataBind.SetBound True, "", "    "
'    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
'    Call SetReDataCombo(oForm, sQry, oCombo1)
'    oForm.Items("CLTCOD").DisplayDesc = True
    
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo1 = Nothing
    Set oCombo2 = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
    
ZPY421_CreateItems_Error:
    Set oCheck = Nothing
    Set oEdit = Nothing
    Set oCombo1 = Nothing
    Set oCombo2 = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set optBtn = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "ZPY421_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub ZPY421_EnableMenus()
    
    On Error GoTo ZPY421_EnableMenus_Error

    Call oForm.EnableMenu("1283", True)     '//제거
    Call oForm.EnableMenu("1284", False)    '//취소
    Call oForm.EnableMenu("1293", True)     '//행삭제
    
    Exit Sub
    
ZPY421_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "ZPY421_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub ZPY421_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo ZPY421_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call ZPY421_FormItemEnabled
        Call ZPY421_AddMatrixRow
    Else
        oForm.Mode = fm_FIND_MODE
        Call ZPY421_FormItemEnabled
        oForm.Items("Code").Specific.Value = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
ZPY421_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "ZPY421_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub ZPY421_FormItemEnabled()
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim i           As Long
    Dim sQry        As String
    
    On Error GoTo ZPY421_FormItemEnabled_Error
    
    
    Call oForm.Freeze(True)
    If (oForm.Mode = fm_ADD_MODE) Then
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
'
'        '//기본사항 - 부서 (사업장에 따른 부서변경)
'        Set oCombo = oForm.Items("DptStr").Specific
'
'        If oCombo.ValidValues.Count > 0 Then
'            For i = oCombo.ValidValues.Count - 1 To 0 Step -1
'                oCombo.ValidValues.Remove i, psk_Index
'            Next i
'            oCombo.ValidValues.Add "", ""
'            oCombo.Select 0, psk_Index
'        End If
'
'        If oForm.Items("CLTCOD").Specific.Value <> "" Then
'            sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
'            sQry = sQry & " WHERE Code = '1' AND U_Char2 = '" & oForm.Items.Item("CLTCOD").Specific.Value & "'"
'            sQry = sQry & " ORDER BY U_Code"
'            Call SetReDataCombo(oForm, sQry, oCombo)
'        End If
        
        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", False)    '//문서추가
        
    ElseIf (oForm.Mode = fm_FIND_MODE) Then
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        
'        '//기본사항 - 부서 (사업장에 따른 부서변경)
'        Set oCombo = oForm.Items("TeamCode").Specific
'
'        If oCombo.ValidValues.Count > 0 Then
'            For i = oCombo.ValidValues.Count - 1 To 0 Step -1
'                oCombo.ValidValues.Remove i, psk_Index
'            Next i
'            oCombo.ValidValues.Add "", ""
'            oCombo.Select 0, psk_Index
'        End If
'
'        If oDS_PH_PY001A.GetValue("U_CLTCOD", 0) <> "" Then
'            sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
'            sQry = sQry & " WHERE Code = '1' AND U_Char2 = '" & oForm.Items.Item("CLTCOD").Specific.Value & "'"
'            sQry = sQry & " ORDER BY U_Code"
'            Call SetReDataCombo(oForm, sQry, oCombo)
'        End If
        
        Call oForm.EnableMenu("1281", False)    '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가
    ElseIf (oForm.Mode = fm_OK_MODE) Then
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        
        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가

    End If
    Call oForm.Freeze(False)
    Exit Sub
    
ZPY421_FormItemEnabled_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "ZPY421_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Public Sub Raise_FormItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
    Dim sQry        As String
    Dim i           As Long
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oRecordSet  As SAPbobsCOM.Recordset
    Dim MstInfo         As ZPAY_g_EmpID
    
    On Error GoTo Raise_FormItemEvent_Error
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Select Case pval.EventType
        Case et_ITEM_PRESSED: '//1
            
            If pval.BeforeAction = True Then
                If pval.ItemUid = "CBtn1" Then   '/ ChooseBtn사원리스트
                    
                    oForm.Items("MSTCOD").CLICK ct_Regular
                    Sbo_Application.ActivateMenuItem ("7425")
                    BubbleEvent = False
                ElseIf pval.ItemUid = "1" And oForm.Mode = fm_ADD_MODE Or oForm.Mode = fm_UPDATE_MODE Then
                    If ZPY421_DataValidCheck = False Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                    If File_Create = False Then
                        BubbleEvent = False
                        Exit Sub
                    Else
                        BubbleEvent = False
                        oForm.Mode = fm_OK_MODE
                    End If
                ElseIf pval.ItemUid = "Btn1" Then
                   oFilePath = ZP_Form.vbGetBrowseDirectory(ZP_Form)
                   oForm.DataSources.UserDataSources("Path").ValueEx = oFilePath
                   BubbleEvent = False
                   Exit Sub
               End If
            ElseIf pval.BeforeAction = False Then
            
            End If
     
'----------------------------------------------------------
        Case et_KEY_DOWN: '//2
            If pval.BeforeAction = True And pval.ItemUid = "MSTCOD" And pval.CharPressed = 9 Then
                If Trim$(oForm.Items("MSTCOD").Specific.String) <> "" And _
                    MDC_SetMod.Value_ChkYn("[@PH_PY001A]", "Code", "'" & Trim$(oForm.Items("MSTCOD").Specific.String) & "'", "") = True Then
                    Sbo_Application.StatusBar.SetText "사원번호를 확인하여 주십시오.", bmt_Short, smt_Error
                    BubbleEvent = False
                End If
            End If
'----------------------------------------------------------
        Case et_GOT_FOCUS: '//3
            Select Case pval.ItemUid
            Case "Mat1", "Grid1"
                If pval.Row > 0 Then
                    oLastItemUID = pval.ItemUid
                    oLastColUID = pval.ColUID
                    oLastColRow = pval.Row
                End If
            Case Else
                oLastItemUID = pval.ItemUid
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
                    If pval.ItemUid = "CLTCOD" Then
                        
                        '//기본사항 - 부서 (사업장에 따른 부서변경)
                        Set oCombo = oForm.Items("DptStr").Specific
                        
                        If oCombo.ValidValues.Count > 0 Then
                            For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                                oCombo.ValidValues.Remove i, psk_Index
                            Next i
                            oCombo.ValidValues.Add "", ""
                            oCombo.Select 0, psk_Index
                        End If
                        
                        sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
                        sQry = sQry & " WHERE Code = '1' AND U_Char2 = '" & oForm.Items.Item("CLTCOD").Specific.Value & "'"
                        sQry = sQry & " ORDER BY U_Code"
                        Call SetReDataCombo(oForm, sQry, oCombo)
                        oCombo.ValidValues.Add "%", "전체"
                        oCombo.Select "%", psk_ByValue
                        oForm.Items("DptStr").DisplayDesc = True
                        
                        '//기본사항 - 부서 (사업장에 따른 부서변경)
                        Set oCombo = oForm.Items("DptEnd").Specific
                        
                        If oCombo.ValidValues.Count > 0 Then
                            For i = oCombo.ValidValues.Count - 1 To 0 Step -1
                                oCombo.ValidValues.Remove i, psk_Index
                            Next i
                            oCombo.ValidValues.Add "", ""
                            oCombo.Select 0, psk_Index
                        End If
                        
                        sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] "
                        sQry = sQry & " WHERE Code = '1' AND U_Char2 = '" & oForm.Items.Item("CLTCOD").Specific.Value & "'"
                        sQry = sQry & " ORDER BY U_Code"
                        Call SetReDataCombo(oForm, sQry, oCombo)
                        oCombo.ValidValues.Add "%", "전체"
                        oCombo.Select "%", psk_ByValue
                        oForm.Items("DptEnd").DisplayDesc = True
                        
                        
                    End If
                End If
            End If
            Call oForm.Freeze(False)
'----------------------------------------------------------
        Case et_CLICK: '//6
            If pval.BeforeAction = True And pval.ItemUid <> "1000001" And pval.ItemUid <> "2" Then
                If oLastItemUID = "STRDAT" Then     '/정산년도
                    If Trim$(oForm.Items(oLastItemUID).Specific.Value) <> "" Then
                        If MDC_SetMod.ChkYearMonth(Trim$(CStr((oForm.Items(oLastItemUID).Specific.Value)))) = False Then
                            oForm.Items(oLastItemUID).Update
                            Sbo_Application.StatusBar.SetText "정산년도를 확인하여 주십시오.", bmt_Short, smt_Error
                            BubbleEvent = False
                        End If
                    End If
                ElseIf oLastItemUID = "MSTCOD" Then
                    If Trim$(oForm.Items(oLastItemUID).Specific.String) <> "" And _
                    MDC_SetMod.Value_ChkYn("[@PH_PY001A]", "Code", "'" & Trim$(oForm.Items(oLastItemUID).Specific.String) & "'", "") = True Then
                        oForm.Items(oLastItemUID).Update
                        Sbo_Application.StatusBar.SetText "사원번호를 확인하여 주십시오.", bmt_Short, smt_Error
                        BubbleEvent = False
                    End If
                End If
            End If
'----------------------------------------------------------
        Case et_DOUBLE_CLICK: '//7

'----------------------------------------------------------
        Case et_MATRIX_LINK_PRESSED '//8

'----------------------------------------------------------
        Case et_MATRIX_COLLAPSE_PRESSED '//9
'----------------------------------------------------------
        Case et_VALIDATE: '//10
            If pval.BeforeAction = False And pval.ItemChanged = True Then
                If pval.ItemUid = "STRDAT" Then
                    'oJsnYear = Mid$(oForm.Items(oUID).Specific.String, 1, 4)
                    oJsnYear = Left$(oForm.DataSources.UserDataSources("STRDAT").ValueEx, 4)
                    oForm.DataSources.UserDataSources("ENDDAT").ValueEx = oJsnYear & "1231"
                    oForm.Items("ENDDAT").Update
                End If
                
                If pval.ItemUid = "ENDDAT" Then
                
                End If
                
                If pval.ItemUid = "MSTCOD" Then
                    If oForm.Items("MSTCOD").Specific.String = "" Then
                        oForm.Items("MSTCOD").Specific.String = ""
                        oForm.DataSources.UserDataSources("MSTNAM").ValueEx = ""
                        oForm.DataSources.UserDataSources("EmpID").ValueEx = ""
                    Else
                        oForm.Items("MSTCOD").Specific.String = UCase$(oForm.Items("MSTCOD").Specific.String)
                        MstInfo = MDC_SetMod.Get_EmpID_InFo(oForm.Items("MSTCOD").Specific.String)
                        oForm.DataSources.UserDataSources("MSTNAM").ValueEx = MstInfo.MSTNAM
                        oForm.DataSources.UserDataSources("EmpID").ValueEx = MstInfo.EmpID
                    End If
                    oForm.Items("MSTNAM").Update
                    oForm.Items("EmpID").Update
                End If
            
            End If
'----------------------------------------------------------
        Case et_MATRIX_LOAD: '//11
'            If pval.BeforeAction = True Then
'            ElseIf pval.BeforeAction = False Then
'                oMat1.LoadFromDataSource
'
'                Call ZPY421_FormItemEnabled
'                Call ZPY421_AddMatrixRow
'
'            End If
'----------------------------------------------------------
        Case et_DATASOURCE_LOAD '//12
           
'----------------------------------------------------------
        Case et_FORM_LOAD: '//16
            
'----------------------------------------------------------
        Case et_FORM_UNLOAD: '//17
            'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
            '컬렉션에서 삭제및 모든 메모리 제거
            'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
            If pval.BeforeAction = True Then
            ElseIf pval.BeforeAction = False Then
                RemoveForms oFormUniqueID
                Set oForm = Nothing
'                Set oDS_ZPY421A = Nothing
'                Set oDS_ZPY421B = Nothing
'
                Set oMat1 = Nothing
                
            End If
'----------------------------------------------------------
        Case et_FORM_ACTIVATE: '//18
            
'----------------------------------------------------------
        Case et_FORM_DEACTIVATE: '//19
            
'----------------------------------------------------------
        Case et_FORM_CLOSE '//20
            
'----------------------------------------------------------
        Case et_FORM_RESIZE '//21
'            If pval.BeforeAction = True Then
'
'            ElseIf pval.BeforeAction = False Then
'
'            End If
'----------------------------------------------------------
        Case et_FORM_KEY_DOWN '//22
            
'----------------------------------------------------------
        Case et_FORM_MENU_HILIGHT '//23
            
'----------------------------------------------------------
        Case et_CHOOSE_FROM_LIST '//27
'            If pval.BeforeAction = True Then
'
'            ElseIf pval.Before_Action = False Then
'
'            End If
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
        End Select
    ElseIf (pval.BeforeAction = False) Then
        Select Case pval.MenuUID
            Case "1283":
                oForm.Mode = fm_ADD_MODE
                Call ZPY421_FormItemEnabled
                Call ZPY421_AddMatrixRow
            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//문서찾기
                Call ZPY421_FormItemEnabled
                Call ZPY421_AddMatrixRow
                oForm.Items("Code").CLICK ct_Regular
            Case "1282":    '//문서추가
                Call ZPY421_FormItemEnabled
                Call ZPY421_AddMatrixRow
            Case "1288", "1289", "1290", "1291":
                Call ZPY421_FormItemEnabled
            Case "1293" '// 행삭제
'                '// [MAT1 용]
'                 If oMat1.RowCount <> oMat1.VisualRowCount Then
'                    oMat1.FlushToDataSource
'
'                    While (i <= oDS_ZPY421B.Size - 1)
'                        If oDS_ZPY421B.GetValue("U_FILD01", i) = "" Then
'                            oDS_ZPY421B.RemoveRecord (i)
'                            i = 0
'                        Else
'                            i = i + 1
'                        End If
'                    Wend
'
'                    For i = 0 To oDS_ZPY421B.Size
'                        Call oDS_ZPY421B.setValue("U_LineNum", i, i + 1)
'                    Next i
'
'                    oMat1.LoadFromDataSource
'                End If
'                Call ZPY421_AddMatrixRow
        End Select
    End If
    Call oForm.Freeze(False)
    Exit Sub
Raise_FormMenuEvent_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_FormMenuEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
    Select Case pval.ItemUid
    Case "Mat1"
        If pval.Row > 0 Then
            oLastItemUID = pval.ItemUid
            oLastColUID = pval.ColUID
            oLastColRow = pval.Row
        End If
    Case Else
        oLastItemUID = pval.ItemUid
        oLastColUID = ""
        oLastColRow = 0
    End Select
    Exit Sub
Raise_RightClickEvent_Error:

    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub ZPY421_AddMatrixRow()
    Dim oRow    As Long
    
    On Error GoTo ZPY421_AddMatrixRow_Error
    
    Call oForm.Freeze(True)
    
'    '//[Mat1 용]
'    oMat1.FlushToDataSource
'    oRow = oMat1.VisualRowCount
'
'    If oMat1.VisualRowCount > 0 Then
'        If Trim(oDS_ZPY421B.GetValue("U_FILD01", oRow - 1)) <> "" Then
'            If oDS_ZPY421B.Size <= oMat1.VisualRowCount Then
'                oDS_ZPY421B.InsertRecord (oRow)
'            End If
'            oDS_ZPY421B.Offset = oRow
'            oDS_ZPY421B.setValue "U_LineNum", oRow, oRow + 1
'            oDS_ZPY421B.setValue "U_FILD01", oRow, ""
'            oDS_ZPY421B.setValue "U_FILD02", oRow, ""
'            oDS_ZPY421B.setValue "U_FILD03", oRow, 0
'            oMat1.LoadFromDataSource
'        Else
'            oDS_ZPY421B.Offset = oRow - 1
'            oDS_ZPY421B.setValue "U_LineNum", oRow - 1, oRow
'            oDS_ZPY421B.setValue "U_FILD01", oRow - 1, ""
'            oDS_ZPY421B.setValue "U_FILD02", oRow - 1, ""
'            oDS_ZPY421B.setValue "U_FILD03", oRow - 1, 0
'            oMat1.LoadFromDataSource
'        End If
'    ElseIf oMat1.VisualRowCount = 0 Then
'        oDS_ZPY421B.Offset = oRow
'        oDS_ZPY421B.setValue "U_LineNum", oRow, oRow + 1
'        oDS_ZPY421B.setValue "U_FILD01", oRow, ""
'        oDS_ZPY421B.setValue "U_FILD02", oRow, ""
'        oDS_ZPY421B.setValue "U_FILD03", oRow, 0
'        oMat1.LoadFromDataSource
'    End If
    
    Call oForm.Freeze(False)
    Exit Sub
ZPY421_AddMatrixRow_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "ZPY421_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub ZPY421_FormClear()
On Error GoTo ZPY421_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'ZPY421'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.Value = 1
    Else
        oForm.Items("DocEntry").Specific.Value = DocEntry
    End If
    Exit Sub
ZPY421_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "ZPY421_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function ZPY421_DataValidCheck() As Boolean
On Error GoTo ZPY421_DataValidCheck_Error
    
    Dim i       As Long
    Dim sQry    As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    ZPY421_DataValidCheck = False
    
    If Trim$(oForm.Items("STRDAT").Specific.String) = "" Or Trim$(oForm.Items("ENDDAT").Specific.String) = "" Then
        Sbo_Application.SetStatusBarMessage "대상일자를 확인하여 주십시오.", bmt_Short, True
        Exit Function
    ElseIf oForm.Items("JSNTYP").Specific.Selected Is Nothing Then
        Sbo_Application.SetStatusBarMessage "기간 코드는 필수입니다. 선택하여 주십시오.", bmt_Short, True
        Exit Function
    ElseIf Trim$(oForm.Items("PRTDAT").Specific.String) = "" Then
        Sbo_Application.SetStatusBarMessage "제출일자는 필수입니다. 입력하여 주십시오.", bmt_Short, True
        Exit Function
    ElseIf oForm.Items("ENDDAT").Specific.String < oForm.Items("STRDAT").Specific.String Then
        Sbo_Application.SetStatusBarMessage "대상기간종료일자가 시작일자보다 작습니다. 확인하여 주십시오.", bmt_Short, True
        Exit Function
    End If
    
    DPTSTR = oForm.Items("DptStr").Specific.Selected.Value
    If DPTSTR = "%" Then DPTSTR = "00000001"
    DPTEND = oForm.Items("DptEnd").Specific.Selected.Value
    If DPTEND = "%" Then DPTEND = "ZZZZZZZZ"
    MSTCOD = oForm.Items("MSTCOD").Specific.Value
    If Trim$(MSTCOD) = "" Then MSTCOD = "%"
'    oJsnGbn = oForm.Items("JsnGbn").Specific.Selected.Value
    
    '//----------------------------------------------------------------------------------
    '//필수 체크
    '//----------------------------------------------------------------------------------

'    oMat1.FlushToDataSource
'    '// Matrix 마지막 행 삭제(DB 저장시)
'    If oDS_ZPY421B.Size > 1 Then oDS_ZPY421B.RemoveRecord (oDS_ZPY421B.Size - 1)
'    oMat1.LoadFromDataSource
    
    ZPY421_DataValidCheck = True
    
    Set oRecordSet = Nothing
    Exit Function
    
ZPY421_DataValidCheck_Error:
    Set oRecordSet = Nothing
    ZPY421_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "ZPY421_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function


Function ZPY421_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo ZPY421_Validate_Error
    ZPY421_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@ZPY421A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.Value & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", bmt_Short, True
        ZPY421_Validate = False
        GoTo ZPY421_Validate_Exit
    End If
'
    If ValidateType = "수정" Then

    ElseIf ValidateType = "행삭제" Then

    ElseIf ValidateType = "취소" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
ZPY421_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
ZPY421_Validate_Error:
    ZPY421_Validate = False
    Sbo_Application.SetStatusBarMessage "ZPY421_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function



Private Sub ZPY421_Print_Report01()

    Dim DocNum          As String
    Dim ErrNum          As Integer
    Dim WinTitle        As String
    Dim ReportName      As String
    Dim sQry            As String
    
    Dim BPLID           As String
    Dim ItmBsort        As String
    Dim DocDate         As String
    
    Dim oRecordSet      As SAPbobsCOM.Recordset
    
    On Error GoTo ZPY421_Print_Report01_Error
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '/ ODBC 연결 체크
    If ConnectODBC = False Then
        GoTo ZPY421_Print_Report01_Error
    End If
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/

    WinTitle = "[S142] 발주서"
    ReportName = "S142_1.rpt"
    sQry = "EXEC ZPY421_1 '" & oForm.Items("8").Specific.Value & "'"
    
    '/ Formula 수식필드
    ReDim gRpt_Formula(1)
    ReDim gRpt_Formula_Value(1)
    
    '/ SubReport
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)
    
    '/ Procedure 실행"
    sQry = "EXEC [PS_PP820_01] '" & BPLID & "', '" & ItmBsort & "', '" & DocDate & "'"
    
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
        If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "1", "Y", "V") = False Then
            Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
        End If
    Else
        Sbo_Application.SetStatusBarMessage "조회된 데이터가 없습니다.", bmt_Short, True
    End If
    
    Exit Sub
    
ZPY421_Print_Report01_Error:
    Sbo_Application.SetStatusBarMessage "ZPY421_Print_Report01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub



Private Function File_Create() As Boolean
On Error GoTo Error_Message
    Dim ErrNum          As Integer
    Dim oStr            As String
    Dim sQry            As String
    Dim sRecordset      As SAPbobsCOM.Recordset
    
    Set sRecordset = oCompany.GetBusinessObject(BoRecordset)
    
    ErrNum = 0
    '/ Question
    If Sbo_Application.MessageBox("전산매체신고 파일을 생성하시겠습니까?", 2, "&Yes!", "&No") = 2 Then
        ErrNum = 1
        GoTo Error_Message
    End If
    oMat1.Clear
    MaxRow = 0
    BUSCNT = 0   '/ B레크드 일련번호
    BUSTOT = 0
    
    '/ 파일경로설정
    If oFilePath = "" Then oFilePath = "C:\EOSDATA"
    oFilePath = IIf(Right$(oFilePath, 1) = "\", oFilePath, oFilePath & "\")
    oStr = CreateFolder(Trim$(oFilePath))
    If Trim$(oStr) <> "" Then
        ErrNum = 5
        GoTo Error_Message
    End If
    
    '/ 갑근 제출자(대리인) 레코드
    If File_Create_ARecord = False Then
        ErrNum = 2
        GoTo Error_Message
    End If
    
    Close #1
    Open FILNAM For Output As #1
        '/ A레코드: 갑근 원천징수의무자별 집계 레코드
        Print #1, MDC_SetMod.sStr(Arec.RECGBN) & MDC_SetMod.sStr(Arec.DTAGBN) & MDC_SetMod.sStr(Arec.TAXCOD) & MDC_SetMod.sStr(Arec.PRTDAT) & _
                  MDC_SetMod.sStr(Arec.RPTGBN) & MDC_SetMod.sStr(Arec.TAXAGE) & MDC_SetMod.sStr(Arec.HOMTID) & MDC_SetMod.sStr(Arec.PGMCOD) & _
                  MDC_SetMod.sStr(Arec.BUSNBR) & MDC_SetMod.sStr(Arec.SANGHO) & MDC_SetMod.sStr(Arec.DAMDPT) & MDC_SetMod.sStr(Arec.DAMNAM) & _
                  MDC_SetMod.sStr(Arec.DAMTEL) & MDC_SetMod.sStr(Arec.BUSCNT) & MDC_SetMod.sStr(Arec.HANCOD) & _
                  MDC_SetMod.sStr(Arec.FILLER)
            
        Call Matrix_AddRow("제출자 레코드 생성 완료!", True)
        '/ B레코드: 갑근 집계 레코드 /***********************************************/
        sQry = "SELECT Code, U_TAXCODE, U_BUSNUM, U_CLTNAME, U_COMPRT, U_PERNUM FROM [@PH_PY005A] T0 "
        sQry = sQry & " WHERE U_WCHCLT = '" & oForm.Items("CLTCOD").Specific.Selected.Value & "' ORDER BY CODE"
        sRecordset.DoQuery sQry
        Do Until sRecordset.EOF
            NEWCNT = 0
            CLTCOD = sRecordset.Fields(0).Value
            
            Brec.TAXCOD = sRecordset.Fields("U_TAXCODE").Value
            Brec.BUSNBR = Replace(sRecordset.Fields("U_BUSNUM").Value, "-", "")
            Brec.SANGHO = Trim$(sRecordset.Fields("U_CLTNAME").Value)
            Brec.COMPRT = Trim$(sRecordset.Fields("U_COMPRT").Value)
            Brec.PERNBR = Replace(Trim$(sRecordset.Fields("U_PERNUM").Value), "-", "")
            
            '/ B레코드: 갑근 집계 레코드
            Select Case File_Create_BRecord
            Case 0
                Call Matrix_AddRow(CLTCOD & "- 징수의무자의 집계 레코드 생성 완료!", True)
                
                '/ C레코드: 갑근 주(현)근무처 레코드
                If File_Create_CRecord = False Then
                    ErrNum = 4
                    GoTo Error_Message
                End If
                Call Matrix_AddRow(CLTCOD & "- 징수의무자의 데이터 레코드" & NEWCNT & "건 생성 완료!", True)
            Case 1  '// B레코드가 생성되지 않은 경우 C레코드 생성하지 않고, 건너뜀
            
            Case 2
                ErrNum = 3
                GoTo Error_Message
            End Select
            
            sRecordset.MoveNext
        Loop
        If BUSTOT <> BUSCNT Then   '/ A레코드의 원천의무자수/B레코드의 원천의무자수
            ErrNum = 6
            GoTo Error_Message
        End If
    Close #1
    Set sRecordset = Nothing
    
    oForm.DataSources.UserDataSources("Path").Value = FILNAM
    Sbo_Application.StatusBar.SetText "전산매체수록이 정상적으로 완료되었습니다.", bmt_Short, smt_Success
    File_Create = True
    
    Exit Function
'///////////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set sRecordset = Nothing
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "취소하였습니다.", bmt_Short, smt_Success
    ElseIf ErrNum = 2 Then
        Sbo_Application.StatusBar.SetText "A레코드(갑근 제출자 레코드) 생성 실패.", bmt_Short, smt_Error
    ElseIf ErrNum = 3 Then
        Sbo_Application.StatusBar.SetText "B레코드(갑근 원천징수의무자별 집계 레코드) 생성 실패.", bmt_Short, smt_Error
    ElseIf ErrNum = 4 Then
        Sbo_Application.StatusBar.SetText "C레코드(갑근 주(현)근무처 레코드) 생성 실패.", bmt_Short, smt_Error
    ElseIf ErrNum = 5 Then
        Sbo_Application.StatusBar.SetText "CreateFolder Error : " & oStr, bmt_Short, smt_Error
    ElseIf ErrNum = 6 Then
        Sbo_Application.StatusBar.SetText "B레코드의 징수의무자수(" & CStr(BUSCNT) & ")와 A레코드의 신고의무자수(" & CStr(BUSTOT) & ")가 일치하지 않습니다.", bmt_Short, smt_Error
    Else
        Sbo_Application.StatusBar.SetText "File_Create 실행 중 오류가 발생했습니다." & Space$(10) & Err.Description, bmt_Short, smt_Error
    End If
    File_Create = False
End Function

Private Function File_Create_ARecord() As Boolean
On Error GoTo Error_Message
    Dim ErrNum          As Integer
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim PRTDAT          As String
    Dim BUSNUM          As String
    Dim CheckA          As String
    
    CheckA = False  '/체크필요유무
    ErrNum = 0
    PRTDAT = Mid$(oForm.Items("PRTDAT").Specific.String, 1, 4) & Mid$(oForm.Items("PRTDAT").Specific.String, 6, 2) & Mid$(oForm.Items("PRTDAT").Specific.String, 9, 2)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '/ 신고의무자수
    
    sQry = " SELECT  COUNT(DISTINCT ISNULL(T1.U_CLTCOD, ''))"
    sQry = sQry & " FROM [@PH_PY115A] T0  INNER JOIN [@PH_PY001A] T1 ON T0.U_MSTCOD = T1.Code"
    sQry = sQry & " INNER JOIN [@PS_HR200L] T2 ON T1.U_TeamCode = T2.U_Code AND T2.Code = '1'"
    sQry = sQry & " INNER JOIN [@PH_PY005A] T3 ON ISNULL(T1.U_CLTCOD, '') = T3.Code"
    sQry = sQry & " WHERE   T3.U_WCHCLT = '" & oForm.Items("CLTCOD").Specific.Selected.Value & "'"
    sQry = sQry & " AND     T2.U_Code BETWEEN " & "'" & DPTSTR & "'" & " AND " & "'" & DPTEND & "'"
    sQry = sQry & " AND     T0.U_MSTCOD LIKE N'%'"
    sQry = sQry & " AND     T0.U_RETPAY > 0"
    sQry = sQry & " AND     T0.U_ENDINT >= '" & oForm.Items("STRDAT").Specific.Value & "'"
    sQry = sQry & " AND     T0.U_ENDINT <= '" & oForm.Items("ENDDAT").Specific.Value & "'"

    
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        BUSTOT = oRecordSet.Fields(0).Value
    End If
    If Val(BUSTOT) = 0 Then
        ErrNum = 3
        GoTo Error_Message
    End If
    
    '/ 업체정보
    sQry = "SELECT * FROM [@PH_PY005A] WHERE Code = '" & oForm.Items("CLTCOD").Specific.Selected.Value & "'"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
        ErrNum = 1
        GoTo Error_Message
    Else
     '/ 파일명
        BUSNUM = Replace(oRecordSet.Fields("U_BUSNUM").Value, "-", "")
        If Len(Trim$(BUSNUM)) <> 10 Then
            ErrNum = 2
            GoTo Error_Message
        End If
        'FILNAM = "C:\EOSDATA\E" & Mid$(BUSNUM, 1, 7) & "." & Mid$(BUSNUM, 8, 3)
        'FILNAM = oFilePath & "E" & Mid$(BUSNUM, 1, 7) & "." & Mid$(BUSNUM, 8, 3)
        FILNAM = oFilePath & "EA" & Mid$(BUSNUM, 1, 7) & "." & Mid$(BUSNUM, 8, 3)     '/ 2010년 변경
        
     '/ B Record /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
        BUSNUM = Replace(oRecordSet.Fields("U_BUSNUM").Value, "-", "")
     '/ A Record /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
        Arec.RECGBN = "A"                                                       '/ 레코드 구분
        'Arec.DTAGBN = IIf(Trim$(oJsnGbn) = "1", "22", "25")
        Arec.DTAGBN = "25"                                                      '/ 자료구분
        Arec.TAXCOD = oRecordSet.Fields("U_TAXCODE").Value                      '/ 세무서
        Arec.PRTDAT = PRTDAT                                                    '/ 제출일자
        Arec.RPTGBN = oRecordSet.Fields("U_TaxDGbn").Value                      '/ 제출자
        Arec.TAXAGE = oRecordSet.Fields("U_TaxDCode").Value                     '/ 세무대리인
        Arec.HOMTID = Trim$(oRecordSet.Fields("U_HOMETID").Value)                '/ 홈텍스ID
        Arec.PGMCOD = "9000"                                                    '/ 세무프로그램코드
        Arec.BUSNBR = Replace(oRecordSet.Fields("U_TAXDBUS").Value, "-", "")    '/ 사업자번호
        Arec.SANGHO = Trim$(oRecordSet.Fields("U_TAXDNAM").Value)                '/ 상호
        Arec.DAMDPT = Trim$(oRecordSet.Fields("U_CHGDPT").Value)                 '/ 담당부서명
        Arec.DAMNAM = Trim$(oRecordSet.Fields("U_CHGNAME").Value)                '/ 담당성명
        Arec.DAMTEL = Trim$(oRecordSet.Fields("U_CHGTEL").Value)                 '/ 담당전화번호
        Arec.BUSCNT = Format$(BUSTOT, String$(Len(Arec.BUSCNT), "0"))             '/ 원천징수의무자수
        Arec.HANCOD = "101"                                                     '/ 한글코드 종류
        Arec.FILLER = Space$(Len(Arec.FILLER))
        
     '/ 필수입력 체크
        If Trim$(Arec.TAXCOD) = "" Then
            Call Matrix_AddRow("A레코드:세무서코드가 누락되었습니다. 입력하여 주십시오.", True, True)
            CheckA = True
        End If
        If Trim$(Arec.RPTGBN) = "" Then
            Call Matrix_AddRow("A레코드:제출자구분가 누락되었습니다. 입력하여 주십시오.", True, True)
            CheckA = True
        End If
        If Trim$(Arec.BUSNBR) = "" Then
            Call Matrix_AddRow("A레코드:제출자사업자번호가 누락되었습니다. 입력하여 주십시오.", True, True)
            CheckA = True
        End If
        If Trim$(Arec.SANGHO) = "" Then
            Call Matrix_AddRow("A레코드:제출자상호가 누락되었습니다. 입력하여 주십시오.", True, True)
            CheckA = True
        End If
        If Trim$(Arec.DAMDPT) = "" Then
            Call Matrix_AddRow("A레코드:담당자부서가 누락되었습니다. 입력하여 주십시오.", True, True)
            CheckA = True
        End If
        If Trim$(Arec.DAMNAM) = "" Then
            Call Matrix_AddRow("A레코드:담당자성명이 누락되었습니다. 입력하여 주십시오.", True, True)
            CheckA = True
        End If
        If Trim$(Arec.DAMTEL) = "" Then
            Call Matrix_AddRow("A레코드:담당자전화번호가 누락되었습니다. 입력하여 주십시오.", True, True)
            CheckA = True
        End If
    End If
    
    If CheckA = False Then
        File_Create_ARecord = True
    Else
        File_Create_ARecord = False
    End If
    Set oRecordSet = Nothing
    Exit Function
'///////////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set oRecordSet = Nothing
   
   If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "귀속년도의 자사정보가 존재하지 않습니다. 등록하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 2 Then
        Sbo_Application.StatusBar.SetText "자사정보등록의 사업자번호가 올바르지 않습니다. 확인하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 3 Then
        Sbo_Application.StatusBar.SetText "자사정보등록의 신고의무사업장이 존재하지 않습니다. 확인하여 주십시오.", bmt_Short, smt_Error
    Else
        Call Matrix_AddRow("A레코드오류: " & Err.Description, False, True)
    End If
    File_Create_ARecord = False
End Function

Private Function File_Create_BRecord() As Integer
On Error GoTo Error_Message
    Dim ErrNum          As Integer
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim CheckB          As String
    
    CheckB = False  '/체크필요유무
    ErrNum = 0
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '/ 집계정보
    sQry = " SELECT   ISNULL(COUNT(T0.U_MSTCOD),0) AS SUM_NEWCNT, "
    sQry = sQry & "   SUM(CASE WHEN ISNULL(T0.U_J01NAM,'')='' THEN 0 ELSE 1 END) AS SUM_OLDCNT, "
    sQry = sQry & "   ISNULL(SUM(T0.U_RETPAY),0)  AS SUM_INCOME, "
    sQry = sQry & "   ISNULL(SUM(T0.U_GULGAB),0)  AS SUM_GULGAB, "
    sQry = sQry & "   ISNULL(SUM(T0.U_GULJUM),0)  AS SUM_GULJUM, "
    sQry = sQry & "   0                           AS SUM_GULNON  "
    sQry = sQry & "  FROM [@PH_PY115A] T0  INNER JOIN [@PH_PY001A] T1 ON T0.U_MSTCOD = T1.Code "
    sQry = sQry & "                        INNER JOIN [OUDP] T2 ON T1.Dept = T2.Code                                  "
    sQry = sQry & "  WHERE   ISNULL(T1.U_CLTCOD, '') = '" & Trim$(CLTCOD) & "'"
    sQry = sQry & "  AND     ISNULL(T1.Branch, '') LIKE " & "N'" & oForm.Items("BPLId").Specific.Selected.Value & "'"
    sQry = sQry & "  AND     T1.U_TeamCode BETWEEN " & "'" & DPTSTR & "'" & " AND " & "'" & DPTEND & "'"
    sQry = sQry & "  AND     T0.U_MSTCOD LIKE " & "N'" & Trim$(MSTCOD) & "'"
    sQry = sQry & "  AND     T0.U_RETPAY > 0"
    sQry = sQry & "  AND     T0.U_ENDINT >= '" & oForm.Items("STRDAT").Specific.String & "'"
    sQry = sQry & "  AND     T0.U_ENDINT <= '" & oForm.Items("ENDDAT").Specific.String & "'"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
        ErrNum = 1
        GoTo Error_Message
    ElseIf oRecordSet.Fields(0).Value = 0 Then
        ErrNum = 1
        GoTo Error_Message
    Else
        BUSCNT = BUSCNT + 1
        '/ B Record /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
        Brec.RECGBN = "B"                                                   '/ 레코드구분
        'Brec.DTAGBN = IIf(Trim$(oJsnGbn) = "1", "22", "25")
        Brec.DTAGBN = "25"                                                  '/ 자료구분
        Brec.TAXCOD = Arec.TAXCOD                                           '/ 세무서
        Brec.BUSCNT = Format$(BUSCNT, String$(Len(Brec.BUSCNT), "0"))         '/ 원천징수의무자수 일련번호
        
        Brec.NEWCNT = Format$(oRecordSet.Fields("SUM_NEWCNT").Value, String$(Len(Brec.NEWCNT), "0"))      'C Record수
        Brec.OLDCNT = Format$(oRecordSet.Fields("SUM_OLDCNT").Value, String$(Len(Brec.OLDCNT), "0"))      'D Record수
        Brec.INCOME = Format$(oRecordSet.Fields("SUM_INCOME").Value, String$(Len(Brec.INCOME), "0"))      '소득금액 총액
        Brec.GULGAB = Format$(oRecordSet.Fields("SUM_GULGAB").Value, String$(Len(Brec.GULGAB), "0"))      '결정 소득세
        Brec.GULCOM = Format$(0, String$(Len(Brec.GULCOM), "0"))                                          '공란
        Brec.GULJUM = Format$(oRecordSet.Fields("SUM_GULJUM").Value, String$(Len(Brec.GULJUM), "0"))      '결정 주민세
        Brec.GULNON = Format$(oRecordSet.Fields("SUM_GULNON").Value, String$(Len(Brec.GULNON), "0"))      '결정 농특세
        Brec.GULTOT = Format$(oRecordSet.Fields("SUM_GULGAB").Value + oRecordSet.Fields("SUM_GULJUM").Value + oRecordSet.Fields("SUM_GULNON").Value, String$(Len(Brec.GULTOT), "0"))
        Brec.RNGCOD = oForm.Items("JSNTYP").Specific.Selected.Value                                     '제출대상 기간
        Brec.FILLER = Space$(Len(Brec.FILLER))

        Print #1, MDC_SetMod.sStr(Brec.RECGBN) & MDC_SetMod.sStr(Brec.DTAGBN) & MDC_SetMod.sStr(Brec.TAXCOD) & _
                MDC_SetMod.sStr(Brec.BUSCNT) & MDC_SetMod.sStr(Brec.BUSNBR) & MDC_SetMod.sStr(Brec.SANGHO) & _
                MDC_SetMod.sStr(Brec.COMPRT) & MDC_SetMod.sStr(Brec.PERNBR) & MDC_SetMod.sStr(Brec.NEWCNT) & _
                MDC_SetMod.sStr(Brec.OLDCNT) & MDC_SetMod.sStr(Brec.INCOME) & MDC_SetMod.sStr(Brec.GULGAB) & _
                MDC_SetMod.sStr(Brec.GULCOM) & MDC_SetMod.sStr(Brec.GULJUM) & MDC_SetMod.sStr(Brec.GULNON) & _
                MDC_SetMod.sStr(Brec.GULTOT) & MDC_SetMod.sStr(Brec.RNGCOD) & MDC_SetMod.sStr(Brec.FILLER)
        
       
     '/ 필수입력 체크
        If Trim$(Brec.BUSNBR) = "" Then
            Call Matrix_AddRow("B레코드:사업자번호가 누락되었습니다. 입력하여 주십시오.", True, True)
            CheckB = True
        End If
        If Trim$(Brec.COMPRT) = "" Then
            Call Matrix_AddRow("B레코드:대표자명이 누락되었습니다. 입력하여 주십시오.", True, True)
            CheckB = True
        End If
        If Trim$(Brec.PERNBR) = "" Then
            Call Matrix_AddRow("B레코드:법인(주민)등록번호가 누락되었습니다. 입력하여 주십시오.", True, True)
            CheckB = True
        End If
    End If
    
    If CheckB = False Then
        File_Create_BRecord = 0
    Else
        File_Create_BRecord = 2
    End If
    Set oRecordSet = Nothing
    Exit Function
'///////////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set oRecordSet = Nothing
   
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "집계레코드가 존재하지 않습니다. 등록하여 주십시오.", bmt_Short, smt_Error
        File_Create_BRecord = 1
    Else
        Call Matrix_AddRow("B레코드오류: " & Err.Description, False)
        File_Create_BRecord = 2
    End If
    
End Function

Private Function File_Create_CRecord() As Boolean
On Error GoTo Error_Message
    Dim ErrNum          As Integer
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim CheckC          As String
    CheckC = False  '/체크필요유무
    ErrNum = 0
    C_MSTCOD = ""
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '/ 사원별 정산정보
    sQry = " SELECT  ISNULL(T4.U_DWEGBN, 0)   AS U_DWEGBN,"
    sQry = sQry & "  ISNULL(T1.brthcountr,'') AS U_DWECOD,"
    sQry = sQry & "  ISNULL(T1.citizenshp,'') AS U_INTCOD,"
    sQry = sQry & "  ISNULL(T4.U_FRGTAX,2)   AS U_FRGTAX,"
    sQry = sQry & "  ISNULL(T4.U_INTGBN,1)   AS U_INTGBN,"
    sQry = sQry & "  ISNULL(T1.govID,'')     AS U_PERNBR,"
    sQry = sQry & "  CASE WHEN T0.U_ST2RET IS NULL THEN T0.U_JINDAT ELSE DATEADD(DD,1,T0.U_ST2RET) END AS ST2RET,"
    sQry = sQry & "  T0.*"
    sQry = sQry & " FROM [@PH_PY115A] T0 "
    sQry = sQry & "      INNER JOIN [OHEM] T1 ON T0.U_EmpID = T1.EmpID"
    sQry = sQry & "      INNER JOIN [@PH_PY001A] T4 ON T0.U_MSTCOD = T4.Code"
    sQry = sQry & " WHERE   ISNULL(T1.U_CLTCOD, '') = '" & Trim$(CLTCOD) & "'"
    sQry = sQry & " AND     ISNULL(T1.Branch, '') LIKE " & "N'" & oForm.Items("BPLId").Specific.Selected.Value & "'"
    sQry = sQry & " AND     T4.U_TeamCode BETWEEN " & "'" & DPTSTR & "'" & " And " & "'" & DPTEND & "'"
    sQry = sQry & " AND     T0.U_MSTCOD LIKE " & "N'" & Trim$(MSTCOD) & "'"
    sQry = sQry & " AND     T0.U_RETPAY > 0"
    sQry = sQry & " AND     T0.U_ENDINT >= '" & oForm.Items("STRDAT").Specific.String & "'"
    sQry = sQry & " AND     T0.U_ENDINT <= '" & oForm.Items("ENDDAT").Specific.String & "'"
    sQry = sQry & " ORDER BY T0.DocEntry"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
'        ErrNum = 1
'        GoTo error_Message
    End If
    Do Until oRecordSet.EOF
        NEWCNT = NEWCNT + 1
        C_MSTCOD = oRecordSet.Fields("U_MSTCOD").Value
        '/ C Record /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
        Crec.RECGBN = "C"
        Crec.DTAGBN = "25"
        Crec.TAXCOD = Arec.TAXCOD
        Crec.SQNNBR = Format$(NEWCNT, String$(Len(Crec.SQNNBR), "0"))  '/ 일련번호
        Crec.BUSNBR = Brec.BUSNBR
        Crec.BUSNUM = Brec.BUSNBR
        Crec.SANGHO = Brec.SANGHO
        
        If Trim$(oRecordSet.Fields("U_J01NAM").Value) <> "" Then
            Crec.JONCNT = Format$(1, String$(Len(Crec.JONCNT), "0"))   '/ 종전근무처수
        Else
            Crec.JONCNT = Format$(0, String$(Len(Crec.JONCNT), "0"))   '/ 종전근무처수
        End If
        Crec.DWEGBN = Format$(oRecordSet.Fields("U_DWEGBN").Value, String$(Len(Crec.DWEGBN), "0"))
        Crec.INTGBN = Format$(oRecordSet.Fields("U_INTGBN").Value, String$(Len(Crec.INTGBN), "0"))
        If Crec.DWEGBN <> "1" Then
            Crec.RGNCOD = oRecordSet.Fields("U_DWECOD").Value  '/ 거주지국
        ElseIf Crec.INTGBN <> "1" Then
            Crec.RGNCOD = oRecordSet.Fields("U_INTCOD").Value  '/ 거주지국
        Else
        'If Crec.DWEGBN = "1" And Crec.INTGBN = "1" Then '/ 비거주자인경우 거주지국코드
            Crec.RGNCOD = ""
        End If
        Crec.STRINT = Format$(Mid$(Replace(oRecordSet.Fields("U_STRINT").Value, "-", ""), 1, 8), String$(Len(Crec.STRINT), "0"))
        Crec.ENDINT = Format$(Mid$(Replace(oRecordSet.Fields("U_ENDINT").Value, "-", ""), 1, 8), String$(Len(Crec.ENDINT), "0"))
        Crec.MSTNAM = oRecordSet.Fields("U_MSTNAM").Value
        Crec.PERNBR = Replace(oRecordSet.Fields("U_PERNBR").Value, "-", "")
        Crec.RETRES = Format$(oRecordSet.Fields("U_RETRES").Value, String$(Len(Crec.RETRES), "0"))     '퇴직  사유
        Crec.TJKPAY = Format$(oRecordSet.Fields("U_TJKPAY").Value, String$(Len(Crec.TJKPAY), "0"))
        Crec.SUDAMT = Format$(oRecordSet.Fields("U_SUDAMT").Value, String$(Len(Crec.SUDAMT), "0"))
        Crec.BHMAMT = Format$(oRecordSet.Fields("U_YILPA1").Value, String$(Len(Crec.BHMAMT), "0"))    '/ 퇴직연금일시금
        Crec.BHMAM1 = Format$(oRecordSet.Fields("U_YILPA2").Value, String$(Len(Crec.BHMAM1), "0"))                  '/ 컬럼을 안만들어놔서 일단 0으로 처리
        Crec.TOTAMT = Format$(oRecordSet.Fields("U_TJKPAY").Value + oRecordSet.Fields("U_YILPA1").Value, String$(Len(Crec.TOTAMT), "0"))
        Crec.TOTAM1 = Format$(oRecordSet.Fields("U_SUDAMT").Value + oRecordSet.Fields("U_YILPA2").Value, String$(Len(Crec.TOTAM1), "0"))
        Crec.BTXPAY = Format$(oRecordSet.Fields("U_BTXPAY").Value, String$(Len(Crec.BTXPAY), "0"))

        Crec.MYNACC = Trim$(oRecordSet.Fields("U_MYNACC").Value)
        Crec.MYNTOT = Format$(oRecordSet.Fields("U_MYNTOT").Value, String$(Len(Crec.MYNTOT), "0"))
        Crec.MYNWON = Format$(oRecordSet.Fields("U_MYNWON").Value, String$(Len(Crec.MYNWON), "0"))
        Crec.MYNBUL = Format$(oRecordSet.Fields("U_MYNBUL").Value, String$(Len(Crec.MYNBUL), "0"))
        Crec.MYNGON = Format$(oRecordSet.Fields("U_MYNGON").Value, String$(Len(Crec.MYNGON), "0"))
        Crec.MYNYIL = Format$(oRecordSet.Fields("U_YILPAY").Value, String$(Len(Crec.MYNYIL), "0"))
        
        Crec.JYNACC = Trim$(oRecordSet.Fields("U_JYNACC").Value)
        Crec.JYNTOT = Format$(oRecordSet.Fields("U_JYNTOT").Value, String$(Len(Crec.JYNTOT), "0"))
        Crec.JYNWON = Format$(oRecordSet.Fields("U_JYNWON").Value, String$(Len(Crec.JYNWON), "0"))
        Crec.JYNBUL = Format$(oRecordSet.Fields("U_JYNBUL").Value, String$(Len(Crec.JYNBUL), "0"))
        Crec.JYNGON = Format$(oRecordSet.Fields("U_JYNGON").Value, String$(Len(Crec.JYNGON), "0"))
        Crec.JYITOT = Format$(oRecordSet.Fields("U_JYIL01").Value + oRecordSet.Fields("U_JYIL02").Value, String$(Len(Crec.JYITOT), "0"))
        
        Crec.SH1JIG = Format$(oRecordSet.Fields("U_SH1JIG").Value, String$(Len(Crec.SH1JIG), "0"))
        Crec.SH3JIG = Format$(oRecordSet.Fields("U_SH3JIG").Value, String$(Len(Crec.SH3JIG), "0"))
        Crec.SH1IYO = Format$(oRecordSet.Fields("U_SH1IYO").Value, String$(Len(Crec.SH1IYO), "0"))    '// 컬럼만 만들어놓고 화면에는 아직 없어용
        Crec.SH3IYO = Format$(oRecordSet.Fields("U_SH3IYO").Value, String$(Len(Crec.SH3IYO), "0"))    '// 컬럼만 만들어놓고 화면에는 아직 없어용
        Crec.SH1GIS = Format$(0, String$(Len(Crec.SH1GIS), "0"))  '// 올해는 컬럼을 안만들어놔서 일단 0으로 처리.
        Crec.SH3GIS = Format$(0, String$(Len(Crec.SH3GIS), "0"))  '// 올해는 컬럼을 안만들어놔서 일단 0으로 처리.
        If Val(Crec.SH1JIG) + Val(Crec.SH3JIG) > 0 Then
            Crec.SH1TIL = Format$(oRecordSet.Fields("U_SH1TIL").Value, String$(Len(Crec.SH1TIL), "0"))
            Crec.SH1SUR = Format$(oRecordSet.Fields("U_SH1SUR").Value, String$(Len(Crec.SH1SUR), "0"))
            Crec.SH1GON = Format$(oRecordSet.Fields("U_SH1GON").Value, String$(Len(Crec.SH1GON), "0"))
            Crec.SH1GWA = Format$(oRecordSet.Fields("U_SH1GWA").Value, String$(Len(Crec.SH1GWA), "0"))
            Crec.SH1YAG = Format$(oRecordSet.Fields("U_SH1YAG").Value, String$(Len(Crec.SH1YAG), "0"))
            Crec.SH1YAS = Format$(oRecordSet.Fields("U_SH1YAS").Value, String$(Len(Crec.SH1YAS), "0"))
        Else
            Crec.SH1TIL = Format$(0, String$(Len(Crec.SH1TIL), "0"))
            Crec.SH1SUR = Format$(0, String$(Len(Crec.SH1SUR), "0"))
            Crec.SH1GON = Format$(0, String$(Len(Crec.SH1GON), "0"))
            Crec.SH1GWA = Format$(0, String$(Len(Crec.SH1GWA), "0"))
            Crec.SH1YAG = Format$(0, String$(Len(Crec.SH1YAG), "0"))
            Crec.SH1YAS = Format$(0, String$(Len(Crec.SH1YAS), "0"))
        End If
        
        Crec.SH2JIG = Format$(oRecordSet.Fields("U_SH2JIG").Value, String$(Len(Crec.SH2JIG), "0"))
        Crec.SH4JIG = Format$(oRecordSet.Fields("U_SH4JIG").Value, String$(Len(Crec.SH4JIG), "0"))
        Crec.SH2IYO = Format$(oRecordSet.Fields("U_SH2IYO").Value, String$(Len(Crec.SH2IYO), "0"))    '// 컬럼만 만들어놓고 화면에는 아직 없어용
        Crec.SH4IYO = Format$(oRecordSet.Fields("U_SH4IYO").Value, String$(Len(Crec.SH4IYO), "0"))    '// 컬럼만 만들어놓고 화면에는 아직 없어용
        Crec.SH2GIS = Format$(0, String$(Len(Crec.SH2GIS), "0"))  '// 올해는 컬럼을 안만들어놔서 일단 0으로 처리.
        Crec.SH4GIS = Format$(0, String$(Len(Crec.SH4GIS), "0"))  '// 올해는 컬럼을 안만들어놔서 일단 0으로 처리.
        If Val(Crec.SH2JIG) + Val(Crec.SH4JIG) > 0 Then
            Crec.SH2TIL = Format$(oRecordSet.Fields("U_SH2TIL").Value, String$(Len(Crec.SH2TIL), "0"))
            Crec.SH2SUR = Format$(oRecordSet.Fields("U_SH2SUR").Value, String$(Len(Crec.SH2SUR), "0"))
            Crec.SH2GON = Format$(oRecordSet.Fields("U_SH2GON").Value, String$(Len(Crec.SH2GON), "0"))
            Crec.SH2GWA = Format$(oRecordSet.Fields("U_SH2GWA").Value, String$(Len(Crec.SH2GWA), "0"))
            Crec.SH2YAG = Format$(oRecordSet.Fields("U_SH2YAG").Value, String$(Len(Crec.SH2YAG), "0"))
            Crec.SH2YAS = Format$(oRecordSet.Fields("U_SH2YAS").Value, String$(Len(Crec.SH2YAS), "0"))
        Else
            Crec.SH2TIL = Format$(0, String$(Len(Crec.SH2TIL), "0"))
            Crec.SH2SUR = Format$(0, String$(Len(Crec.SH2SUR), "0"))
            Crec.SH2GON = Format$(0, String$(Len(Crec.SH2GON), "0"))
            Crec.SH2GWA = Format$(0, String$(Len(Crec.SH2GWA), "0"))
            Crec.SH2YAG = Format$(0, String$(Len(Crec.SH2YAG), "0"))
            Crec.SH2YAS = Format$(0, String$(Len(Crec.SH2YAS), "0"))
        End If
        
        '/ 법정 주현
        Crec.INPDAT = Format$(Replace(oRecordSet.Fields("U_STRRET").Value, "-", ""), String$(Len(Crec.INPDAT), "0"))
        Crec.OUTDAT = Format$(Replace(oRecordSet.Fields("U_ENDRET").Value, "-", ""), String$(Len(Crec.OUTDAT), "0"))
        Crec.GNMMON = Format$(oRecordSet.Fields("U_GNMMON").Value, String$(Len(Crec.GNMMON), "0"))
        Crec.EXPMON = Format$(oRecordSet.Fields("U_EXPMON").Value, String$(Len(Crec.EXPMON), "0"))
        '/ 법정 종전
        Crec.JIPDAT = Format$(Replace(oRecordSet.Fields("U_JINDAT").Value, "-", ""), String$(Len(Crec.JIPDAT), "0"))
        Crec.JOTDAT = Format$(Replace(oRecordSet.Fields("U_JOTDAT").Value, "-", ""), String$(Len(Crec.JOTDAT), "0"))
        Crec.JGMMON = Format$(oRecordSet.Fields("U_GNMDAY").Value, String$(Len(Crec.JGMMON), "0"))  '/ 종전근무월(GNMDAY로대체)
        Crec.JEXMON = Format$(oRecordSet.Fields("U_JEXMON").Value, String$(Len(Crec.JEXMON), "0"))
        Crec.DUPMON = Format$(oRecordSet.Fields("U_JMMMON").Value, String$(Len(Crec.DUPMON), "0"))
        Crec.GNMYER = Format$(oRecordSet.Fields("U_GNMYER").Value, String$(Len(Crec.GNMYER), "0"))
        
        '/ 법정외 주현
        
        If Left(Crec.JOTDAT, 4) = Left(Crec.INPDAT, 4) Then '/ 동일년도 중도정산후 퇴사자일경우 주현입사일대신 정산시작일입력함.
            Crec.FR2DAT = Format$(Replace(oRecordSet.Fields("U_STRRET").Value, "-", ""), String$(Len(Crec.FR2DAT), "0"))
        Else
            Crec.FR2DAT = Format$(Replace(oRecordSet.Fields("U_INPDAT").Value, "-", ""), String$(Len(Crec.FR2DAT), "0"))
        End If
        Crec.TO2DAT = Format$(Replace(oRecordSet.Fields("U_ENDRET").Value, "-", ""), String$(Len(Crec.TO2DAT), "0"))
        If oRecordSet.Fields("U_GN2MON").Value = 0 Or (Left(Crec.JOTDAT, 4) = Left(Crec.INPDAT, 4)) Then    '/법정외근속월수가 없을경우 법정 근속내용삽입
            Crec.GN2MON = Format$(oRecordSet.Fields("U_GNMMON").Value, String$(Len(Crec.GN2MON), "0"))
            Crec.EX2MON = Format$(oRecordSet.Fields("U_EXPMON").Value, String$(Len(Crec.EX2MON), "0"))
            Crec.GN2YER = Format$(oRecordSet.Fields("U_GNMYER").Value, String$(Len(Crec.GNMYER), "0"))
        Else
            Crec.GN2MON = Format$(oRecordSet.Fields("U_GN2MON").Value, String$(Len(Crec.GN2MON), "0"))
            Crec.EX2MON = Format$(oRecordSet.Fields("U_EX2MON").Value, String$(Len(Crec.EX2MON), "0"))
            Crec.GN2YER = Format$(oRecordSet.Fields("U_GN2YER").Value, String$(Len(Crec.GN2YER), "0"))
        End If
        '/ 법정외 종전
        Crec.JI2DAT = Format$(Replace(oRecordSet.Fields("ST2RET").Value, "-", ""), String$(Len(Crec.JI2DAT), "0"))    '전  입사일자
        Crec.JO2DAT = Format$(Replace(oRecordSet.Fields("U_JOTDAT").Value, "-", ""), String$(Len(Crec.JO2DAT), "0"))
        If oRecordSet.Fields("U_JO2MON").Value = 0 Then  '/ 종전 법정외근속월수가 없을경우 법정 근속내용삽입
            Crec.JO2MON = Format$(oRecordSet.Fields("U_GNMDAY").Value, String$(Len(Crec.JO2MON), "0")) '    근속월수
            Crec.JE2MON = Format$(oRecordSet.Fields("U_JEXMON").Value, String$(Len(Crec.JE2MON), "0")) '    제외월수(2007년)
            Crec.JB2MON = Format$(oRecordSet.Fields("U_JMMMON").Value, String$(Len(Crec.JB2MON), "0"))
        Else
            Crec.JO2MON = Format$(oRecordSet.Fields("U_JO2MON").Value, String$(Len(Crec.JO2MON), "0")) '    근속월수
            Crec.JE2MON = Format$(oRecordSet.Fields("U_JE2MON").Value, String$(Len(Crec.JE2MON), "0")) '    제외월수(2007년)
            Crec.JB2MON = Format$(oRecordSet.Fields("U_JB2MON").Value, String$(Len(Crec.JB2MON), "0"))
        End If
        
        '/ 법정
        Crec.JS1RET = Format$(Val(oRecordSet.Fields("U_RETPAY").Value) - Val(oRecordSet.Fields("U_SH2SUR").Value), String$(Len(Crec.JS1RET), "0"))
        Crec.JS1GON = Format$(Val(oRecordSet.Fields("U_RETGON").Value) - Val(oRecordSet.Fields("U_SH2GON").Value), String$(Len(Crec.JS1GON), "0"))
        Crec.JS1STD = Format$(Val(oRecordSet.Fields("U_TAXSTD").Value) - Val(oRecordSet.Fields("U_SH2GWA").Value), String$(Len(Crec.JS1STD), "0"))
        Crec.JS1YSD = Format$(Val(oRecordSet.Fields("U_YTXSTD").Value) - Val(oRecordSet.Fields("U_SH2YAG").Value), String$(Len(Crec.JS1YSD), "0"))
        Crec.JS1YST = Format$(Val(oRecordSet.Fields("U_YSANTX").Value) - Val(oRecordSet.Fields("U_SH2YAS").Value), String$(Len(Crec.JS1YST), "0"))
        Crec.JS1SAN = Format$(Val(oRecordSet.Fields("U_SANTAX").Value) - Val(oRecordSet.Fields("U_JS2SAN").Value), String$(Len(Crec.JS1SAN), "0"))
        Crec.JS1FRN = Format$(Val(oRecordSet.Fields("U_TAXGON").Value) - Val(oRecordSet.Fields("U_JS2TAX").Value), String$(Len(Crec.JS1FRN), "0"))  '/2008년추가 세액(외국납부)공제
        '/ 법정외
        Crec.JS2RET = Format$(oRecordSet.Fields("U_SH2SUR").Value, String$(Len(Crec.JS2RET), "0"))
        Crec.JS2GON = Format$(oRecordSet.Fields("U_SH2GON").Value, String$(Len(Crec.JS2GON), "0"))
        Crec.JS2STD = Format$(oRecordSet.Fields("U_SH2GWA").Value, String$(Len(Crec.JS2STD), "0"))
        Crec.JS2YSD = Format$(oRecordSet.Fields("U_SH2YAG").Value, String$(Len(Crec.JS2YSD), "0"))
        Crec.JS2YST = Format$(oRecordSet.Fields("U_SH2YAS").Value, String$(Len(Crec.JS2YST), "0"))
        Crec.JS2SAN = Format$(oRecordSet.Fields("U_JS2SAN").Value, String$(Len(Crec.JS2SAN), "0"))
        Crec.JS2FRN = Format$(oRecordSet.Fields("U_JS2TAX").Value, String$(Len(Crec.JS2FRN), "0"))  '/2008년추가 세액(외국납부)공제
        '/ 계
        Crec.RETPAY = Format$(oRecordSet.Fields("U_RETPAY").Value, String$(Len(Crec.RETPAY), "0"))
        Crec.RETGON = Format$(oRecordSet.Fields("U_RETGON").Value, String$(Len(Crec.RETGON), "0"))
        Crec.TAXSTD = Format$(oRecordSet.Fields("U_TAXSTD").Value, String$(Len(Crec.TAXSTD), "0"))
        Crec.YTXSTD = Format$(oRecordSet.Fields("U_YTXSTD").Value, String$(Len(Crec.YTXSTD), "0"))
        Crec.YSANTX = Format$(oRecordSet.Fields("U_YSANTX").Value, String$(Len(Crec.YSANTX), "0"))
        Crec.SANTAX = Format$(oRecordSet.Fields("U_SANTAX").Value, String$(Len(Crec.SANTAX), "0"))
        Crec.TAXGON = Format$(oRecordSet.Fields("U_TAXGON").Value, String$(Len(Crec.TAXGON), "0"))  '/2008년추가 세액(외국납부)공제
        '/ 결정
        Crec.GULGAB = Format$(oRecordSet.Fields("U_GULGAB").Value, String$(Len(Crec.GULGAB), "0"))
        Crec.GULJUM = Format$(oRecordSet.Fields("U_GULJUM").Value, String$(Len(Crec.GULJUM), "0"))
        Crec.GULNON = Format$(0, String$(Len(Crec.GULNON), "0"))
        Crec.GULTOT = Format$(oRecordSet.Fields("U_GULGAB").Value + oRecordSet.Fields("U_GULJUM").Value, String$(Len(Crec.GULTOT), "0"))
        '/ 종전
        Crec.JONGAB = Format$(oRecordSet.Fields("U_JONGAB").Value, String$(Len(Crec.JONGAB), "0"))
        Crec.JONJUM = Format$(oRecordSet.Fields("U_JONJUM").Value, String$(Len(Crec.JONJUM), "0"))
        Crec.JONNON = Format$(0, String$(Len(Crec.JONNON), "0"))
        Crec.JONTOT = Format$(oRecordSet.Fields("U_JONGAB").Value + oRecordSet.Fields("U_JONJUM").Value, String$(Len(Crec.JONTOT), "0"))
        Crec.FILLER = Space$(Len(Crec.FILLER))
        
        Print #1, MDC_SetMod.sStr(Crec.RECGBN) & MDC_SetMod.sStr(Crec.DTAGBN) & MDC_SetMod.sStr(Crec.TAXCOD) & MDC_SetMod.sStr(Crec.SQNNBR) & MDC_SetMod.sStr(Crec.BUSNBR) & MDC_SetMod.sStr(Crec.BUSNUM) & _
                MDC_SetMod.sStr(Crec.SANGHO) & MDC_SetMod.sStr(Crec.JONCNT) & MDC_SetMod.sStr(Crec.DWEGBN) & MDC_SetMod.sStr(Crec.RGNCOD) & _
                MDC_SetMod.sStr(Crec.STRINT) & MDC_SetMod.sStr(Crec.ENDINT) & MDC_SetMod.sStr(Crec.MSTNAM) & MDC_SetMod.sStr(Crec.INTGBN) & MDC_SetMod.sStr(Crec.PERNBR) & MDC_SetMod.sStr(Crec.RETRES) & _
                MDC_SetMod.sStr(Crec.TJKPAY) & MDC_SetMod.sStr(Crec.SUDAMT) & MDC_SetMod.sStr(Crec.BHMAMT) & MDC_SetMod.sStr(Crec.BHMAM1) & MDC_SetMod.sStr(Crec.TOTAMT) & MDC_SetMod.sStr(Crec.TOTAM1) & MDC_SetMod.sStr(Crec.BTXPAY) & _
                MDC_SetMod.sStr(Crec.MYNACC) & MDC_SetMod.sStr(Crec.MYNTOT) & MDC_SetMod.sStr(Crec.MYNWON) & MDC_SetMod.sStr(Crec.MYNBUL) & MDC_SetMod.sStr(Crec.MYNGON) & MDC_SetMod.sStr(Crec.MYNYIL) & _
                MDC_SetMod.sStr(Crec.JYNACC) & MDC_SetMod.sStr(Crec.JYNTOT) & MDC_SetMod.sStr(Crec.JYNWON) & MDC_SetMod.sStr(Crec.JYNBUL) & MDC_SetMod.sStr(Crec.JYNGON) & MDC_SetMod.sStr(Crec.JYITOT) & _
                MDC_SetMod.sStr(Crec.SH1JIG) & MDC_SetMod.sStr(Crec.SH3JIG) & MDC_SetMod.sStr(Crec.SH1IYO) & MDC_SetMod.sStr(Crec.SH3IYO) & MDC_SetMod.sStr(Crec.SH1GIS) & MDC_SetMod.sStr(Crec.SH3GIS) & _
                MDC_SetMod.sStr(Crec.SH1TIL) & MDC_SetMod.sStr(Crec.SH1SUR) & MDC_SetMod.sStr(Crec.SH1GON) & MDC_SetMod.sStr(Crec.SH1GWA) & MDC_SetMod.sStr(Crec.SH1YAG) & MDC_SetMod.sStr(Crec.SH1YAS) & _
                MDC_SetMod.sStr(Crec.SH2JIG) & MDC_SetMod.sStr(Crec.SH4JIG) & MDC_SetMod.sStr(Crec.SH2IYO) & MDC_SetMod.sStr(Crec.SH4IYO) & MDC_SetMod.sStr(Crec.SH2GIS) & MDC_SetMod.sStr(Crec.SH4GIS) & _
                MDC_SetMod.sStr(Crec.SH2TIL) & MDC_SetMod.sStr(Crec.SH2SUR) & MDC_SetMod.sStr(Crec.SH2GON) & MDC_SetMod.sStr(Crec.SH2GWA) & MDC_SetMod.sStr(Crec.SH2YAG) & MDC_SetMod.sStr(Crec.SH2YAS) & _
                MDC_SetMod.sStr(Crec.INPDAT) & MDC_SetMod.sStr(Crec.OUTDAT) & MDC_SetMod.sStr(Crec.GNMMON) & MDC_SetMod.sStr(Crec.EXPMON) & MDC_SetMod.sStr(Crec.JIPDAT) & MDC_SetMod.sStr(Crec.JOTDAT) & _
                MDC_SetMod.sStr(Crec.JGMMON) & MDC_SetMod.sStr(Crec.JEXMON) & MDC_SetMod.sStr(Crec.DUPMON) & MDC_SetMod.sStr(Crec.GNMYER) & _
                MDC_SetMod.sStr(Crec.FR2DAT) & MDC_SetMod.sStr(Crec.TO2DAT) & MDC_SetMod.sStr(Crec.GN2MON) & MDC_SetMod.sStr(Crec.EX2MON) & MDC_SetMod.sStr(Crec.JI2DAT) & MDC_SetMod.sStr(Crec.JO2DAT) & _
                MDC_SetMod.sStr(Crec.JO2MON) & MDC_SetMod.sStr(Crec.JE2MON) & MDC_SetMod.sStr(Crec.JB2MON) & MDC_SetMod.sStr(Crec.GN2YER) & _
                MDC_SetMod.sStr(Crec.JS1RET) & MDC_SetMod.sStr(Crec.JS1GON) & MDC_SetMod.sStr(Crec.JS1STD) & MDC_SetMod.sStr(Crec.JS1YSD) & MDC_SetMod.sStr(Crec.JS1YST) & MDC_SetMod.sStr(Crec.JS1SAN) & MDC_SetMod.sStr(Crec.JS1FRN) & _
                MDC_SetMod.sStr(Crec.JS2RET) & MDC_SetMod.sStr(Crec.JS2GON) & MDC_SetMod.sStr(Crec.JS2STD) & MDC_SetMod.sStr(Crec.JS2YSD) & MDC_SetMod.sStr(Crec.JS2YST) & MDC_SetMod.sStr(Crec.JS2SAN) & MDC_SetMod.sStr(Crec.JS2FRN) & _
                MDC_SetMod.sStr(Crec.RETPAY) & MDC_SetMod.sStr(Crec.RETGON) & MDC_SetMod.sStr(Crec.TAXSTD) & MDC_SetMod.sStr(Crec.YTXSTD) & MDC_SetMod.sStr(Crec.YSANTX) & MDC_SetMod.sStr(Crec.SANTAX) & MDC_SetMod.sStr(Crec.TAXGON) & _
                MDC_SetMod.sStr(Crec.GULGAB) & MDC_SetMod.sStr(Crec.GULJUM) & MDC_SetMod.sStr(Crec.GULNON) & MDC_SetMod.sStr(Crec.GULTOT) & MDC_SetMod.sStr(Crec.JONGAB) & MDC_SetMod.sStr(Crec.JONJUM) & MDC_SetMod.sStr(Crec.JONNON) & _
                MDC_SetMod.sStr(Crec.JONTOT) & MDC_SetMod.sStr(Crec.FILLER)
        
        Call Matrix_AddRow("C레코드:" & C_MSTCOD & Trim$(Crec.MSTNAM) & " 생성 완료.", True)
     '/ 필수입력 체크
        If Crec.DWEGBN = "0" Then
            Call Matrix_AddRow("C레코드:" & C_MSTCOD & Trim$(Crec.MSTNAM) & "-거주자구분이 누락되었습니다. 확인하여 주십시오.", False, True)
            CheckC = True
        End If
        If Crec.DWEGBN = "2" And Trim$(Crec.RGNCOD) = "" Then
            Call Matrix_AddRow("C레코드:" & C_MSTCOD & Trim$(Crec.MSTNAM) & "-비거주자일경우 거주지국코드는 필수입니다. 확인하여 주십시오.", False, True)
            CheckC = True
        End If
        If Crec.INTGBN = "1" And Len(Trim$(Crec.PERNBR)) <> 13 Then
            Call Matrix_AddRow("C레코드:" & C_MSTCOD & Trim$(Crec.MSTNAM) & "-주민등록번호를 확인하여 주십시오.", False, True)
            CheckC = True
        End If
        If Crec.RETRES = "0" Or Trim$(Crec.RETRES) = "" Then
            Call Matrix_AddRow("C레코드:" & C_MSTCOD & Trim$(Crec.MSTNAM) & "-퇴직사유항목을 입력하여 주십시오.", False, True)
            CheckC = True
        End If
        
        '/ D레코드: 종전근무처 레코드
        If Val(Crec.JONCNT) > 0 Then
            Drec.RECGBN = "D"
            'DREC.DTAGBN = IIf(Trim$(oJsnGbn) = "1", "22", "25")
            Drec.DTAGBN = "25"
            Drec.TAXCOD = Arec.TAXCOD
            Drec.SQNNBR = Format$(NEWCNT, String$(Len(Drec.SQNNBR), "0"))
            
            Drec.BUSNBR = Brec.BUSNBR
            Drec.FILLD1 = Space$(Len(Drec.FILLD1))
            Drec.PERNBR = Crec.PERNBR
            
            Drec.JONNAM = Trim$(oRecordSet.Fields("U_J01NAM").Value)
            Drec.JONNBR = Replace(oRecordSet.Fields("U_J01NBR").Value, "-", "")
            Drec.RETAMT = Format$(oRecordSet.Fields("U_JRET01").Value, String$(Len(Drec.RETAMT), "0"))
            Drec.SUDAMT = Format$(oRecordSet.Fields("U_JSUD01").Value, String$(Len(Drec.SUDAMT), "0"))
            Drec.BHMAMT = Format$(oRecordSet.Fields("U_JBHM01").Value, String$(Len(Drec.BHMAMT), "0"))
            Drec.BHMAM1 = Format$(0, String$(Len(Drec.BHMAM1), "0"))          '/ 컬럼을 안만들어놔서 일단 0으로 처리
            Drec.TOTAMT = Format$(oRecordSet.Fields("U_JTOT01").Value, String$(Len(Drec.TOTAMT), "0"))
            Drec.TOTAM1 = Format$(oRecordSet.Fields("U_JADD01").Value, String$(Len(Drec.TOTAM1), "0"))
            Drec.BTXP01 = Format$(oRecordSet.Fields("U_BTXP01").Value, String$(Len(Drec.BTXP01), "0"))
            Drec.JONCNT = Format$(Crec.JONCNT, String$(Len(Drec.JONCNT), "0"))
            Drec.FILLER = Space$(Len(Drec.FILLER))
            
            Print #1, MDC_SetMod.sStr(Drec.RECGBN) & MDC_SetMod.sStr(Drec.DTAGBN) & MDC_SetMod.sStr(Drec.TAXCOD) & MDC_SetMod.sStr(Drec.SQNNBR) & _
                      MDC_SetMod.sStr(Drec.BUSNBR) & MDC_SetMod.sStr(Drec.FILLD1) & MDC_SetMod.sStr(Drec.PERNBR) & MDC_SetMod.sStr(Drec.JONNAM) & MDC_SetMod.sStr(Drec.JONNBR) & _
                      MDC_SetMod.sStr(Drec.RETAMT) & MDC_SetMod.sStr(Drec.SUDAMT) & MDC_SetMod.sStr(Drec.BHMAMT) & MDC_SetMod.sStr(Drec.BHMAM1) & MDC_SetMod.sStr(Drec.TOTAMT) & _
                      MDC_SetMod.sStr(Drec.TOTAM1) & MDC_SetMod.sStr(Drec.BTXP01) & _
                      MDC_SetMod.sStr(Drec.JONCNT) & MDC_SetMod.sStr(Drec.FILLER)
            
            
            '/ 필수입력 체크
               If Trim$(Drec.JONNAM) = "" Then
                   Call Matrix_AddRow("D레코드:종전근무처 법인명(상호)가 누락되었습니다. 입력하여 주십시오.", True, True)
                   CheckC = True
               End If
               If Trim$(Drec.JONNBR) = "" Then
                   Call Matrix_AddRow("D레코드:종전근무처 사업자번호가 누락되었습니다. 입력하여 주십시오.", True, True)
                   CheckC = True
               End If
        End If
        
        oRecordSet.MoveNext
    Loop
        
        
    If CheckC = False Then
        File_Create_CRecord = True
    Else
        File_Create_CRecord = False
    End If
    Set oRecordSet = Nothing
    Exit Function
'///////////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set oRecordSet = Nothing
   
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "주(현)근무처 레코드가 존재하지 않습니다. 등록하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 2 Then
        Sbo_Application.StatusBar.SetText "D레코드(종전근무처 레코드) 생성 실패.", bmt_Short, smt_Error
    Else
        Call Matrix_AddRow("C레코드오류: " & Err.Description, False)
    End If
    File_Create_CRecord = False
End Function

Private Sub Matrix_AddRow(ByVal MatrixMsg As String, Optional Insert_YN As Boolean, Optional MatrixErr As Boolean)
    If MatrixErr = True Then
        oForm.DataSources.UserDataSources("Col0").Value = "??"
    Else
        oForm.DataSources.UserDataSources("Col0").Value = ""
    End If
    oForm.DataSources.UserDataSources("Col1").Value = MatrixMsg
    If Insert_YN = True Then
        oMat1.AddRow
        MaxRow = MaxRow + 1
    End If
    oMat1.SetLineData MaxRow
End Sub
