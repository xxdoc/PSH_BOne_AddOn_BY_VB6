Attribute VB_Name = "MDC_Globals"
Option Explicit

'//전역변수
Public oCompany                 As SAPbobsCOM.Company
Public Sbo_Application          As SAPbouiCOM.Application

Public oProgBar             As SAPbouiCOM.ProgressBar

Public FormCurrentCount         As Long         '//현재 폼의총갯수
Public FormTotalCount           As Long         '//생성한 폼의총갯수
Public ClassList                As Collection   '//컬렉션 개체
Public FormTypeListCount        As Long         '//FormType 객체수
Public FormTypeList             As Collection   '//FormType 객체

Public oForm_ActiveItem         As String
Public oForm_ActiveRow          As Integer

'//Path/Srf/Rpt 패스
Public SP_XMLPath               As String   '//XML메뉴경로
Public SP_Path                  As String   '//PathINI경로
Public SP_Screen                As String   '//스크린폴더위치
Public SP_Report                As String   '//레포트폴더위치

'//ODBC
Public SP_ODBC_YN               As String
Public SP_ODBC_Name             As String
Public SP_ODBC_DBName           As String
Public SP_ODBC_ID               As String
Public SP_ODBC_PW               As String

'//Network Connection
Public SP_NETWORK_YN            As String
Public SP_NETWORK_DRIVE         As String
Public SP_NETWORK_PATH          As String
Public SP_NETWORK_ID            As String
Public SP_NETWORK_PW            As String

'//Cr부분
Public ZG_CRWDSN                As String
Public g_ERPDMS                 As ADODB.Connection
Public g_ADORS1                 As ADODB.Recordset
Public g_ADORS2                 As ADODB.Recordset

Public g_CApp                   As CRAXDDRT.Application
Public g_Report                 As CRAXDDRT.Report
Public g_cFormula               As CRAXDDRT.FormulaFieldDefinition
Public g_GCrview                As Object
Public g_Params                 As CRAXDDRT.ParameterFieldDefinitions
Public g_Param                  As CRAXDDRT.ParameterFieldDefinition

Public g_CrSections             As CRAXDDRT.Sections
Public g_CrSection              As CRAXDDRT.Section
Public g_CrReportObjs           As CRAXDDRT.ReportObjects
Public g_CrSubReportObj         As CRAXDDRT.SubreportObject
Public g_CrSubReport            As CRAXDDRT.Report
Public g_CrDB                   As CRAXDDRT.Database

Public gRpt_Formula()           As String
Public gRpt_Formula_Value()     As String
Public gRpt_Param()             As String
Public gRpt_Param_Value()       As String
Public gRpt_SRptSqry()          As String
Public gRpt_SRptName()          As String
Public gRpt_SFormula()          As String
Public gRpt_SFormula_Value()    As String


'//사원조회 저장용 변수
Public Type ZPAY_g_EmpID
    EmpID     As String      '//사원순번
    MSTCOD    As String      '//사원번호
    MSTNAM    As String      '//사원성명
    TeamCode  As String      '//부서
    RspCode   As String      '//담당
    ClsCode   As String      '//반
    CLTCOD    As String      '//자사코드
    StartDate As String      '//입사일자
    TermDate  As String      '//퇴사일자
    RETDAT    As String      '//퇴직정산일
    BALYMD    As String      '//최종발령일
    BALCOD    As String      '//최종부서
    JIGTYP    As String      '//직종
    Position  As String      '//직위
    JIGCOD    As String      '//직급
    HOBONG    As String      '//호봉
    PAYTYP    As String      '//급여형태
    PAYSEL    As String      '//급여지급일구분
    GONCNT    As Integer     '//공제인원
    DAGYSU    As Integer     '//다자녀추가공제
    STDAMT    As Double      '//기본급
    GBHSEL    As String      '//고용보험여부
    PERNBR    As String      '//주민번호
    Sex       As String      '//성별
    GRPDAT    As String      '//그룹입사일
    ENDRET    As String      '//퇴직중간정산일
End Type

Public M_Used(1 To 4) As Boolean     '/ 1:근태, 2:급상여, 3:퇴직, 4:원천
Public M_DayGNT As Boolean     '/ 일근태사용
Public M_YunGNT As Boolean     '/ 년차사용
Public M_PrsGNT As Boolean     '/ 개인별근태사용
Public M_JsnGIT As Boolean     '/ 정산기타소득사용
Public M_JsnBUS As Boolean     '/ 정산사업소득사용
Public M_JsnEJA As Boolean     '/ 정산이자소득사용
Public M_JsnILY As Boolean     '/ 정산일용직사용


'//사용자구조체
Public Value01 As String
Public Value02 As String
Public Value03 As String
Public Value04 As String
Public Value05 As String
Public Value06 As String
Public Value07 As String
Public Value08 As String
Public Value09 As String
Public Value10 As String
Public Value11 As String
Public Value12 As String
Public Value13 As String
Public Value14 As String
Public Value15 As String
Public Value16 As String
Public Value17 As String
Public Value18 As String
Public Value19 As String
Public Value20 As String

Public oTitleNameCount          As Long

Public ZP_Form              As Form
Public frmRPT_View11        As Form
Public frmRPT_View12        As Form
Public frmRPT_View13        As Form


Public ZPAY_GBL_GNSYER     As Integer          '근 속  년 수
Public ZPAY_GBL_GNSMON     As Integer          '       월 수
Public ZPAY_GBL_GNSDAY     As Integer          '       일 수
Public ZPAY_GBL_GNMYER     As Integer          '근 무  년 수
Public ZPAY_GBL_GNMMON     As Integer          '       월 수
Public ZPAY_GBL_GNMDAY     As Integer          '       일 수

Public ZPAY_GBL_JSNYER     As String * 4       '정산년도
