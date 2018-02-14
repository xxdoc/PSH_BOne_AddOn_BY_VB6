Attribute VB_Name = "MDC_Globals"
Option Explicit

'//��������
Public oCompany                 As SAPbobsCOM.Company
Public Sbo_Application          As SAPbouiCOM.Application

Public oProgBar             As SAPbouiCOM.ProgressBar

Public FormCurrentCount         As Long         '//���� �����Ѱ���
Public FormTotalCount           As Long         '//������ �����Ѱ���
Public ClassList                As Collection   '//�÷��� ��ü
Public FormTypeListCount        As Long         '//FormType ��ü��
Public FormTypeList             As Collection   '//FormType ��ü

Public oForm_ActiveItem         As String
Public oForm_ActiveRow          As Integer

'//Path/Srf/Rpt �н�
Public SP_XMLPath               As String   '//XML�޴����
Public SP_Path                  As String   '//PathINI���
Public SP_Screen                As String   '//��ũ��������ġ
Public SP_Report                As String   '//����Ʈ������ġ

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

'//Cr�κ�
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


'//�����ȸ ����� ����
Public Type ZPAY_g_EmpID
    EmpID     As String      '//�������
    MSTCOD    As String      '//�����ȣ
    MSTNAM    As String      '//�������
    TeamCode  As String      '//�μ�
    RspCode   As String      '//���
    ClsCode   As String      '//��
    CLTCOD    As String      '//�ڻ��ڵ�
    StartDate As String      '//�Ի�����
    TermDate  As String      '//�������
    RETDAT    As String      '//����������
    BALYMD    As String      '//�����߷���
    BALCOD    As String      '//�����μ�
    JIGTYP    As String      '//����
    Position  As String      '//����
    JIGCOD    As String      '//����
    HOBONG    As String      '//ȣ��
    PAYTYP    As String      '//�޿�����
    PAYSEL    As String      '//�޿������ϱ���
    GONCNT    As Integer     '//�����ο�
    DAGYSU    As Integer     '//���ڳ��߰�����
    STDAMT    As Double      '//�⺻��
    GBHSEL    As String      '//��뺸�迩��
    PERNBR    As String      '//�ֹι�ȣ
    Sex       As String      '//����
    GRPDAT    As String      '//�׷��Ի���
    ENDRET    As String      '//�����߰�������
End Type

Public M_Used(1 To 4) As Boolean     '/ 1:����, 2:�޻�, 3:����, 4:��õ
Public M_DayGNT As Boolean     '/ �ϱ��»��
Public M_YunGNT As Boolean     '/ �������
Public M_PrsGNT As Boolean     '/ ���κ����»��
Public M_JsnGIT As Boolean     '/ �����Ÿ�ҵ���
Public M_JsnBUS As Boolean     '/ �������ҵ���
Public M_JsnEJA As Boolean     '/ �������ڼҵ���
Public M_JsnILY As Boolean     '/ �����Ͽ������


'//����ڱ���ü
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


Public ZPAY_GBL_GNSYER     As Integer          '�� ��  �� ��
Public ZPAY_GBL_GNSMON     As Integer          '       �� ��
Public ZPAY_GBL_GNSDAY     As Integer          '       �� ��
Public ZPAY_GBL_GNMYER     As Integer          '�� ��  �� ��
Public ZPAY_GBL_GNMMON     As Integer          '       �� ��
Public ZPAY_GBL_GNMDAY     As Integer          '       �� ��

Public ZPAY_GBL_JSNYER     As String * 4       '����⵵
