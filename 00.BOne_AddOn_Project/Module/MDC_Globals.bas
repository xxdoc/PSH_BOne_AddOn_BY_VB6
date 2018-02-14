Attribute VB_Name = "MDC_Globals"
Option Explicit
Public ProgramType              As String
Public Z_Language               As String
'
'//로그인부분

Public gParam_ODBC              As String
Public gParam_DataBase          As String
Public gParam_DBID              As String
Public gParam_DBPW              As String
Public gParam_Server            As String


Public ZG_CRWDSN                As String

'//Cr부분
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

'//RFC
Public oSapConnection01         As Object

