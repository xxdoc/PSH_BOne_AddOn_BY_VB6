VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ZZMDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1
Private WithEvents Sbo_Application      As SAPbouiCOM.Application   'SBO application
Attribute Sbo_Application.VB_VarHelpID = -1

Private Sub Class_Initialize()
    
    On Error GoTo Class_Initialize_Error
    
    Call Initialize
    Call SetApplication
    Call ConnectToSBO
    Call Execute
    Call InitializeSetting
    Call ConnectShareFolder
    Call AddMenuXml
    Call DoSomething
    
    Sbo_Application.StatusBar.SetText "애드온이 초기화되었습니다.", bmt_Short, smt_Success
    
    Exit Sub
    
Class_Initialize_Error:
    Sbo_Application.SetStatusBarMessage "Class_Initialize_Error : " & Err.Number & "-" & Err.Description, bmt_Short, True
End Sub

Private Sub Initialize()
    MDC_Globals.FormCurrentCount = 0
    MDC_Globals.FormTotalCount = 0
    Set MDC_Globals.ClassList = New Collection
    MDC_Globals.FormTypeListCount = 0
    Set MDC_Globals.FormTypeList = New Collection
End Sub

Private Sub SetApplication()

    Dim SboGuiApi As New SAPbouiCOM.SboGuiApi
    Dim ConnectionString As String
    
    On Error GoTo SetApplication_Error
    
    ConnectionString = Command
    If Trim(ConnectionString) = "" Then
        ConnectionString = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056"
    End If
    Call SboGuiApi.Connect(ConnectionString)
    Set Sbo_Application = SboGuiApi.GetApplication()
    Set MDC_Globals.Sbo_Application = Sbo_Application
    Sbo_Application.SetStatusBarMessage "Addon을 시작합니다.", bmt_Short, False
    Exit Sub
    
SetApplication_Error:
    Call MsgBox("SAP BusinessOne 에 접속할수 없습니다. 응용프로그램이 정상적으로 동작하는지 확인하시기 바랍니다.", vbOKOnly, "SAP Business One")
    End
End Sub

Public Sub ConnectToSBO()
    
    On Error GoTo ConnectToSBO_Error
    
    Set oCompany = Sbo_Application.Company.GetDICompany()
    If oCompany.Connected = False Then
        Sbo_Application.MessageBox "DB에 접속을 실패하였습니다."
        End
    End If
    Sbo_Application.SetStatusBarMessage "DB에 접속하였습니다.", bmt_Short, False
    Exit Sub
    
ConnectToSBO_Error:
    Call MsgBox("DB에 접속을 실패하였습니다.", vbOKOnly, "SAP Business One")
    End
End Sub

Public Sub InitializeSetting()
    Dim sQry        As String
    Dim oRecordSet  As SAPbobsCOM.Recordset

    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
'
'    sQry = "SELECT U_Code AS Code, U_Value01 AS Value01, U_Value02 AS Value02, U_Value03 AS Value03, U_Value04 AS Value04, U_Value05 AS Value05 "
'    sQry = sQry & "FROM [@PH_SY001L] "
'    sQry = sQry & "WHERE Code = 'CSY001'"
'
'    oRecordset.DoQuery sQry
'
'    If oRecordset.RecordCount > 0 Then
'
'        Do While Not oRecordset.EOF
'            Select Case UCase(Trim(oRecordset.Fields("Code").Value))
'            Case UCase("PATH")
'                If Trim(oRecordset.Fields("Value01").Value) = "Y" Then
'                    MDC_Globals.SP_Path = Trim(oRecordset.Fields("Value02").Value)
'                    MDC_Globals.SP_Screen = Trim(oRecordset.Fields("Value03").Value)
'                    MDC_Globals.SP_Report = Trim(oRecordset.Fields("Value04").Value)
'                Else
'                    MDC_Globals.SP_XMLPath = App.Path
'                    MDC_Globals.SP_Path = App.Path & "\PathINI"
'                    MDC_Globals.SP_Screen = "HR_Screen"
'                    MDC_Globals.SP_Report = "HR_Report"
'                End If
'            Case UCase("ODBC")
'                MDC_Globals.SP_ODBC_YN = Trim(oRecordset.Fields("Value01").Value)
'                MDC_Globals.SP_ODBC_Name = Trim(oRecordset.Fields("Value02").Value)
'                MDC_Globals.SP_ODBC_DBName = oCompany.CompanyDB
'                MDC_Globals.SP_ODBC_ID = Trim(oRecordset.Fields("Value04").Value)
'                MDC_Globals.SP_ODBC_PW = Trim(oRecordset.Fields("Value05").Value)
'            Case UCase("NETWORK")
'                MDC_Globals.SP_NETWORK_YN = Trim(oRecordset.Fields("Value01").Value)
'                MDC_Globals.SP_NETWORK_DRIVE = Trim(oRecordset.Fields("Value02").Value)
'                MDC_Globals.SP_NETWORK_PATH = Trim(oRecordset.Fields("Value03").Value)
'                MDC_Globals.SP_NETWORK_ID = Trim(oRecordset.Fields("Value04").Value)
'                MDC_Globals.SP_NETWORK_PW = Trim(oRecordset.Fields("Value05").Value)
'
'            End Select
'            oRecordset.MoveNext
'        Loop
'    End If

    '//풍산용
    sQry = "SELECT PARAM01 AS PARAM01,PARAM02 AS PARAM02,PARAM03 AS PARAM03,PARAM04 AS PARAM04,PARAM05 AS PARAM05,PARAM06 AS PARAM06,PARAM07 AS PARAM07,PARAM08 AS PARAM08"
    sQry = sQry & " FROM PROFILE "
    sQry = sQry & " WHERE TYPE = 'SERVERINFO'"

    oRecordSet.DoQuery sQry

    If oRecordSet.RecordCount > 0 Then

    '// PATH
        MDC_Globals.SP_XMLPath = App.Path
        MDC_Globals.SP_Path = App.Path & "\PathINI"
        MDC_Globals.SP_Screen = "HR_Screen"
        MDC_Globals.SP_Report = "HR_Report"
    '// ODBC
'        MDC_Globals.SP_ODBC_YN = Trim(oRecordset.Fields("Value01").Value)
        MDC_Globals.SP_ODBC_Name = "MDCERP"
        MDC_Globals.SP_ODBC_DBName = oCompany.CompanyDB
        MDC_Globals.SP_ODBC_ID = Trim(oRecordSet.Fields("PARAM07").VALUE)
        MDC_Globals.SP_ODBC_PW = Trim(oRecordSet.Fields("PARAM08").VALUE)
'    '//NETWORK
'        MDC_Globals.SP_NETWORK_YN = Trim(oRecordset.Fields("Value01").Value)
'        MDC_Globals.SP_NETWORK_DRIVE = Trim(oRecordset.Fields("Value02").Value)
'        MDC_Globals.SP_NETWORK_PATH = Trim(oRecordset.Fields("Value03").Value)
'        MDC_Globals.SP_NETWORK_ID = Trim(oRecordset.Fields("Value04").Value)
'        MDC_Globals.SP_NETWORK_PW = Trim(oRecordset.Fields("Value05").Value)
    End If
        
    Set oRecordSet = Nothing
End Sub

Private Sub ConnectShareFolder()

    On Error GoTo ConnectShareFolder_Error
    
    If (MDC_Globals.SP_NETWORK_YN = "Y") Then
        Call SetDrive(MDC_Globals.SP_NETWORK_DRIVE, MDC_Globals.SP_NETWORK_PATH, MDC_Globals.SP_NETWORK_ID, MDC_Globals.SP_NETWORK_PW) '//127.0.0.1\B1_SHR\PathINI
    End If
    Exit Sub
    
ConnectShareFolder_Error:
    Sbo_Application.SetStatusBarMessage "공유폴더 연결에 실패 하였습니다 - " & Err.Description, bmt_Short, True
End Sub

Private Sub AddMenuXml()
    
    Dim FileName    As String
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    
    FileName = "Menu_KOR.xml"
    
    oXmlDoc.Load (MDC_Globals.SP_XMLPath & "\" & FileName)
    Sbo_Application.LoadBatchActions oXmlDoc.xml
End Sub

Private Sub TerminateApplication()
    Dim WS
    Dim oDrives
    Dim i           As Long
    Dim oTempClass  As Variant

    If ClassList.Count > 0 Then
        For i = 0 To ClassList.Count - 1
            Set oTempClass = ClassList.Item(i)
            ClassList.Remove i
            Set oTempClass = Nothing
        Next i
    End If
    
    oCompany.Disconnect
    
    Set oCompany = Nothing
    Set g_ERPDMS = Nothing
    End
End Sub

Private Sub Sbo_Application_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes)
    Select Case EventType
    
        Case aet_CompanyChanged:
'            Call TerminateApplication
            End
        Case aet_LanguageChanged:
            Call AddMenuXml
            
        Case aet_ServerTerminition:
'            Call TerminateApplication
            End
        Case aet_ShutDown:
'            Call TerminateApplication
            End
    End Select
End Sub

Private Sub Sbo_Application_ItemEvent(ByVal FormUID As String, pval As SAPbouiCOM.ItemEvent, BubbleEvent As Boolean)

    Dim oTempClass          As Variant
    
    On Error GoTo Sbo_Application_ItemEvent_Error
    
    Call CreateSYSTEMForm(pval)
    
    If Left(pval.FormUID, 2) <> "F_" Then
        If ValidateForm(pval.FormTypeEx) Then
            Set oTempClass = ClassList.Item(FormUID)
            If (oTempClass.oForm Is Nothing) Then
                Exit Sub
            Else
                Call oTempClass.Raise_FormItemEvent(FormUID, pval, BubbleEvent)
            End If
        End If
    ElseIf Left(pval.FormUID, 2) = "F_" Then
        If ValidateForm(pval.FormTypeEx) Then
            Set oTempClass = ClassList.Item(FormUID)
            Call oTempClass.Raise_FormItemEvent(FormUID, pval, BubbleEvent)
        End If
    End If
    Exit Sub
    
Sbo_Application_ItemEvent_Error:
    If Err.Number = 5 Then
        '//LoadForm이 호출되기전에 발생하는 이벤트
        '//EX> S41폼 호출전 GOT_FOCUS 이벤트발생
    Else
        Sbo_Application.SetStatusBarMessage "SBO_Application_ItemEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    End If
End Sub

Private Sub Sbo_Application_MenuEvent(pval As SAPbouiCOM.IMenuEvent, BubbleEvent As Boolean)

    Dim oTempClass          As Variant
    
    On Error GoTo Sbo_Application_MenuEvent_Error
    
    If pval.BeforeAction = True Then
        Call CreateUSERForm(pval)
    End If

    Dim FormUID As String
    FormUID = Sbo_Application.Forms.ActiveForm.uniqueID
    
    If Left(FormUID, 2) <> "F_" Then
        If ValidateForm(Sbo_Application.Forms.ActiveForm.TypeEx) Then
            Set oTempClass = ClassList.Item(FormUID)
            If (oTempClass.oForm Is Nothing) Then
                Exit Sub
            Else
                Call oTempClass.Raise_FormMenuEvent(FormUID, pval, BubbleEvent)
            End If
        End If
    ElseIf Left(FormUID, 2) = "F_" Then
        If ValidateForm(Sbo_Application.Forms.ActiveForm.TypeEx) Then
            Set oTempClass = ClassList.Item(FormUID)
            Call oTempClass.Raise_FormMenuEvent(FormUID, pval, BubbleEvent)
        End If
    End If
    Exit Sub
    
Sbo_Application_MenuEvent_Error:
    If Err.Number = 5 Then
        '//모두닫기메뉴실행시 이전에존재하는 폼에의해발생
    ElseIf Err.Number = -3012 Then
        '//주메뉴닫기시발생하는오류
    Else
        Sbo_Application.SetStatusBarMessage "Sbo_Application_MenuEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    End If
End Sub

Private Sub Sbo_Application_FormDataEvent(BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, BubbleEvent As Boolean)

    Dim oTempClass As Variant
    Dim FormUID As String
    
    On Error GoTo Sbo_Application_FormDataEvent_Error

    FormUID = BusinessObjectInfo.FormUID
    
    If Left(FormUID, 2) <> "F_" Then
        If ValidateForm(BusinessObjectInfo.FormTypeEx) Then
            Set oTempClass = ClassList.Item(FormUID)
            If (oTempClass.oForm Is Nothing) Then
                Exit Sub
            Else
                Call oTempClass.Raise_FormDataEvent(FormUID, BusinessObjectInfo, BubbleEvent)
            End If
        End If
    ElseIf Left(FormUID, 2) = "F_" Then
        If ValidateForm(BusinessObjectInfo.FormTypeEx) Then
            Set oTempClass = ClassList.Item(FormUID)
            Call oTempClass.Raise_FormDataEvent(FormUID, BusinessObjectInfo, BubbleEvent)
        End If
    End If
    Exit Sub
    
Sbo_Application_FormDataEvent_Error:
    
    Sbo_Application.SetStatusBarMessage "Sbo_Application_FormDataEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True

End Sub

Private Sub Sbo_Application_RightClickEvent(eventInfo As SAPbouiCOM.IContextMenuInfo, BubbleEvent As Boolean)

    Dim FormUID As String
    Dim oTempClass As Variant
    
    On Error GoTo Sbo_Application_RightClickEvent_Error
    
    FormUID = eventInfo.FormUID
    If Left(FormUID, 2) <> "F_" Then
        If ValidateForm(Sbo_Application.Forms.Item(eventInfo.FormUID).TypeEx) Then
            Set oTempClass = ClassList.Item(FormUID)
            If (oTempClass.oForm Is Nothing) Then
                Exit Sub
            Else
                Call oTempClass.Raise_RightClickEvent(FormUID, eventInfo, BubbleEvent)
            End If
        End If
    ElseIf Left(FormUID, 2) = "F_" Then
        If ValidateForm(Sbo_Application.Forms.Item(eventInfo.FormUID).TypeEx) Then
            Set oTempClass = ClassList.Item(FormUID)
            Call oTempClass.Raise_RightClickEvent(FormUID, eventInfo, BubbleEvent)
        End If
    End If
    Exit Sub
    
Sbo_Application_RightClickEvent_Error:
    Sbo_Application.SetStatusBarMessage "Sbo_Application_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function ValidateForm(ByVal FormType As String) As Boolean

    Dim i As Long
    
    On Error GoTo ValidateForm_Error:
    
    For i = 1 To FormTypeListCount
        If (FormTypeList.Item(i) = FormType) Then
            ValidateForm = True
            Exit Function
        End If
    Next
    ValidateForm = False
    Exit Function
    
ValidateForm_Error:
    Sbo_Application.SetStatusBarMessage "ValidateForm_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub CreateUSERForm(ByRef pval As SAPbouiCOM.IMenuEvent)

    Dim oTempClass          As Variant
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    
    
    On Error GoTo CreateUSERForm_Error
    
    If (pval.BeforeAction = True) Then
    
        sQry = " EXEC Z_PS_FormCount '" & MDC_PS_Common.User_MSTCOD & "','" & pval.MenuUID & "'"
        Call oRecordSet.DoQuery(sQry)
        
        
         Select Case pval.MenuUID
         
        '---------------------------------------------------------------------------------------
        '// 권한
        '---------------------------------------------------------------------------------------
        Case "PH_PY000":  Set oTempClass = New PH_PY000: Call oTempClass.LoadForm      '//사용자권한관리
        
        '---------------------------------------------------------------------------------------
        '// 인사
        '---------------------------------------------------------------------------------------
        Case "PH_PY001":  Set oTempClass = New PH_PY001: Call oTempClass.LoadForm      '//인사마스터등록
        Case "PH_PY002":  Set oTempClass = New PH_PY002: Call oTempClass.LoadForm      '//근태시간구분등록
        Case "PH_PY003":  Set oTempClass = New PH_PY003: Call oTempClass.LoadForm      '//근태월력설정
        Case "PH_PY004":  Set oTempClass = New PH_PY004: Call oTempClass.LoadForm      '//근무조편성등록
        Case "PH_PY005":  Set oTempClass = New PH_PY005: Call oTempClass.LoadForm      '//사업장정보등록
        Case "PH_PY006":  Set oTempClass = New PH_PY006: Call oTempClass.LoadForm      '//승호작업등록
        Case "PH_PY007":  Set oTempClass = New PH_PY007: Call oTempClass.LoadForm      '//유류단가등록
        Case "PH_PY008":  Set oTempClass = New PH_PY008: Call oTempClass.LoadForm      '//일근태등록
        Case "PH_PY009":  Set oTempClass = New PH_PY009: Call oTempClass.LoadForm      '//기찰기파일upload
        Case "PH_PY010":  Set oTempClass = New PH_PY010: Call oTempClass.LoadForm      '//일일근태처리
        Case "PH_PY011":  Set oTempClass = New PH_PY011: Call oTempClass.LoadForm      '//전문직호칭일괄변경
        Case "PH_PY012":  Set oTempClass = New PH_PY012: Call oTempClass.LoadForm      '//출장등록
        Case "PH_PY013":  Set oTempClass = New PH_PY013: Call oTempClass.LoadForm      '//위해일수 계산
        Case "PH_PY014":  Set oTempClass = New PH_PY014: Call oTempClass.LoadForm      '//위해일수 수정
        Case "PH_PY015":  Set oTempClass = New PH_PY015: Call oTempClass.LoadForm      '//연차적치등록
        Case "PH_PY016":  Set oTempClass = New PH_PY016: Call oTempClass.LoadForm      '//기본업무등록
        Case "PH_PY017":  Set oTempClass = New PH_PY017: Call oTempClass.LoadForm      '//월근태집계처리
        Case "PH_PY018":  Set oTempClass = New PH_PY018: Call oTempClass.LoadForm      '//휴일근무체크(연봉제)
        Case "PH_PY019":  Set oTempClass = New PH_PY019: Call oTempClass.LoadForm      '//반변경등록
        Case "PH_PY020":  Set oTempClass = New PH_PY020: Call oTempClass.LoadForm      '//일근태 업무변경 등록
        Case "PH_PY021":  Set oTempClass = New PH_PY021: Call oTempClass.LoadForm      '//사원비상연락처관리
        Case "PH_PY201":  Set oTempClass = New PH_PY201: Call oTempClass.LoadForm      '//정년임박자 휴가경비 등록
        Case "PH_PY202":  Set oTempClass = New PH_PY202: Call oTempClass.LoadForm      '//정년임박자 휴가경비 등록 조회
        Case "PH_PY203":  Set oTempClass = New PH_PY203: Call oTempClass.LoadForm      '//교육실적등록
        Case "PH_PY204":  Set oTempClass = New PH_PY204: Call oTempClass.LoadForm      '//교육계획등록
        Case "PH_PY205":  Set oTempClass = New PH_PY205: Call oTempClass.LoadForm      '//교육계획VS실적조회
         
        '// 인사 - 리포트
        Case "PH_PY501":  Set oTempClass = New PH_PY501: Call oTempClass.LoadForm      '//여권발급현황
        Case "PH_PY505":  Set oTempClass = New PH_PY505: Call oTempClass.LoadForm      '//입사자대장
        Case "PH_PY510":  Set oTempClass = New PH_PY510: Call oTempClass.LoadForm      '//사원명부
        Case "PH_PY515":  Set oTempClass = New PH_PY515: Call oTempClass.LoadForm      '//재직자사원명부
        Case "PH_PY520":  Set oTempClass = New PH_PY520: Call oTempClass.LoadForm      '//퇴직및퇴직예정자대장
        Case "PH_PY525":  Set oTempClass = New PH_PY525: Call oTempClass.LoadForm      '//학력별인원현황
        Case "PH_PY530":  Set oTempClass = New PH_PY530: Call oTempClass.LoadForm      '//연령별인원현황
        Case "PH_PY535":  Set oTempClass = New PH_PY535: Call oTempClass.LoadForm      '//근속년수별인원현황
        Case "PH_PY540":  Set oTempClass = New PH_PY540: Call oTempClass.LoadForm      '//인원현황(대외용)
        Case "PH_PY545":  Set oTempClass = New PH_PY545: Call oTempClass.LoadForm      '//인원현황(대내용)
        Case "PH_PY550":  Set oTempClass = New PH_PY550: Call oTempClass.LoadForm      '//전체인원현황
        Case "PH_PY555":  Set oTempClass = New PH_PY555: Call oTempClass.LoadForm      '//일일근무자현황
        Case "PH_PY560":  Set oTempClass = New PH_PY560: Call oTempClass.LoadForm      '//일출근현황
        Case "PH_PY565":  Set oTempClass = New PH_PY565: Call oTempClass.LoadForm      '//연장근무자현황
        Case "PH_PY570":  Set oTempClass = New PH_PY570: Call oTempClass.LoadForm      '//연장/휴일근무자현황
        Case "PH_PY575":  Set oTempClass = New PH_PY575: Call oTempClass.LoadForm      '//근태기찰현황황
        Case "PH_PY580":  Set oTempClass = New PH_PY580: Call oTempClass.LoadForm      '//개인별근태월보
        Case "PH_PY585":  Set oTempClass = New PH_PY585: Call oTempClass.LoadForm      '//일일출근기록부
        Case "PH_PY590":  Set oTempClass = New PH_PY590: Call oTempClass.LoadForm      '//기간별근태집계표
        Case "PH_PY595":  Set oTempClass = New PH_PY595: Call oTempClass.LoadForm      '//근속년수현황
        Case "PH_PY600":  Set oTempClass = New PH_PY600: Call oTempClass.LoadForm      '//일자별연장근무현황
        Case "PH_PY605":  Set oTempClass = New PH_PY605: Call oTempClass.LoadForm      '//근속보전휴가발생및사용내역
        Case "PH_PY610":  Set oTempClass = New PH_PY610: Call oTempClass.LoadForm      '//근태구분별사용내역
        Case "PH_PY615":  Set oTempClass = New PH_PY615: Call oTempClass.LoadForm      '//당직근무현황
        Case "PH_PY620":  Set oTempClass = New PH_PY620: Call oTempClass.LoadForm      '//연봉제휴일근무자현황
        Case "PH_PY635":  Set oTempClass = New PH_PY635: Call oTempClass.LoadForm      '//여행,교육자현황
        Case "PH_PY640":  Set oTempClass = New PH_PY640: Call oTempClass.LoadForm      '//국민연금퇴직전환금현황
        Case "PH_PY645":  Set oTempClass = New PH_PY645: Call oTempClass.LoadForm      '//자격수당지급현황
        Case "PH_PY650":  Set oTempClass = New PH_PY650: Call oTempClass.LoadForm      '//노동조합간부현황
        Case "PH_PY655":  Set oTempClass = New PH_PY655: Call oTempClass.LoadForm      '//보훈대상자현황
        Case "PH_PY660":  Set oTempClass = New PH_PY660: Call oTempClass.LoadForm      '//장애근로자현황
        Case "PH_PY665":  Set oTempClass = New PH_PY665: Call oTempClass.LoadForm      '//사원자녀현황
        Case "PH_PY670":  Set oTempClass = New PH_PY670: Call oTempClass.LoadForm      '//개인별차량현황
        Case "PH_PY675":  Set oTempClass = New PH_PY675: Call oTempClass.LoadForm      '//근무편성현황
        Case "PH_PY680":  Set oTempClass = New PH_PY680: Call oTempClass.LoadForm      '//상벌현황
        Case "PH_PY685":  Set oTempClass = New PH_PY685: Call oTempClass.LoadForm      '//포상가급현황
        Case "PH_PY690":  Set oTempClass = New PH_PY690: Call oTempClass.LoadForm      '//생일자현황
        Case "PH_PY695":  Set oTempClass = New PH_PY695: Call oTempClass.LoadForm      '//인사기록카드
        Case "PH_PY705":  Set oTempClass = New PH_PY705: Call oTempClass.LoadForm      '//교통비지급근태확인
        Case "PH_PY677":  Set oTempClass = New PH_PY677: Call oTempClass.LoadForm      '//일일근태이상자조회(2013.03.25 송명규 추가)
        Case "PH_PY676":  Set oTempClass = New PH_PY676: Call oTempClass.LoadForm      '//근태시간내역(2013.06.06 송명규 추가)
        Case "PH_PY860":  Set oTempClass = New PH_PY860: Call oTempClass.LoadForm      '//호봉표조회
        Case "PH_PY503":  Set oTempClass = New PH_PY503: Call oTempClass.LoadForm      '//승진대상자명부
        Case "PH_PY678":  Set oTempClass = New PH_PY678: Call oTempClass.LoadForm      '//당직근무자 일괄 등록(2013.10.30 송명규 추가)
        Case "PH_PY679":  Set oTempClass = New PH_PY679: Call oTempClass.LoadForm      '//개인별 근태집계(기간)
        Case "PH_PY507":  Set oTempClass = New PH_PY507: Call oTempClass.LoadForm      '//휴직자현황
        Case "PH_PY681":  Set oTempClass = New PH_PY681: Call oTempClass.LoadForm      '//비근무일수현황(2014.05.10 송명규 추가)
        Case "PH_PY935":  Set oTempClass = New PH_PY935: Call oTempClass.LoadForm      '//정기승호현황
        Case "PH_PY551":  Set oTempClass = New PH_PY551: Call oTempClass.LoadForm      '//평균인원조회(2015.04.07 송명규 추가)
        Case "PH_PY508":  Set oTempClass = New PH_PY508: Call oTempClass.LoadForm      '//재직증명 등록 및 발급(2015.05.14 송명규 추가)
        Case "PH_PY522":  Set oTempClass = New PH_PY522: Call oTempClass.LoadForm      '//임금피크대상자현황
        Case "PH_PY523":  Set oTempClass = New PH_PY523: Call oTempClass.LoadForm      '//임금피크대상자월별차수현황
        Case "PH_PY524":  Set oTempClass = New PH_PY524: Call oTempClass.LoadForm      '//임금피크대상자월별차수현황
        Case "PH_PY683":  Set oTempClass = New PH_PY683: Call oTempClass.LoadForm      '//교대근무인정현황
        Case "PH_PYA65":  Set oTempClass = New PH_PYA65: Call oTempClass.LoadForm      '//년차현황(집계)
        Case "PH_PY583":  Set oTempClass = New PH_PY583: Call oTempClass.LoadForm      '//근태마감 체크
        

        '---------------------------------------------------------------------------------------
        '// 급여
        '---------------------------------------------------------------------------------------
        Case "PH_PY100":  Set oTempClass = New PH_PY100: Call oTempClass.LoadForm      '//기준세엑설정
        Case "PH_PY101":  Set oTempClass = New PH_PY101: Call oTempClass.LoadForm      '//보험율등록
        Case "PH_PY102":  Set oTempClass = New PH_PY102: Call oTempClass.LoadForm      '//수당항목설정
        Case "PH_PY103":  Set oTempClass = New PH_PY103: Call oTempClass.LoadForm      '//공제항목설정
        Case "PH_PY104":  Set oTempClass = New PH_PY104: Call oTempClass.LoadForm      '//고정수당공제금액일괄등록
        Case "PH_PY105":  Set oTempClass = New PH_PY105: Call oTempClass.LoadForm      '//호봉표등록
        Case "PH_PY106":  Set oTempClass = New PH_PY106: Call oTempClass.LoadForm      '//수당계산식설정
        Case "PH_PY107":  Set oTempClass = New PH_PY107: Call oTempClass.LoadForm      '//급상여기준일설정
        Case "PH_PY108":  Set oTempClass = New PH_PY108: Call oTempClass.LoadForm      '//상여지급율설정
        Case "PH_PY109":  Set oTempClass = New PH_PY109: Call oTempClass.LoadForm      '//급상여변동자료등록
        Case "PH_PY109_1":  Set oTempClass = New PH_PY109_1: Call oTempClass.LoadForm  '//급상여변동자료 수정
        Case "PH_PY110":  Set oTempClass = New PH_PY110: Call oTempClass.LoadForm      '//개인별상여율등록
        Case "PH_PY111":  Set oTempClass = New PH_PY111: Call oTempClass.LoadForm      '//급상여계산
        Case "PH_PY112":  Set oTempClass = New PH_PY112: Call oTempClass.LoadForm      '//급상여자료관리
        Case "PH_PY113":  Set oTempClass = New PH_PY113: Call oTempClass.LoadForm      '//급상여분개자료생성
        Case "PH_PY114":  Set oTempClass = New PH_PY114: Call oTempClass.LoadForm      '//퇴직금기준설정
        Case "PH_PY115":  Set oTempClass = New PH_PY115: Call oTempClass.LoadForm      '//퇴직금계산
        Case "PH_PY116":  Set oTempClass = New PH_PY116: Call oTempClass.LoadForm      '//퇴직금분개자료생성
        Case "PH_PY117":  Set oTempClass = New PH_PY117: Call oTempClass.LoadForm      '//급상여마감작업
        Case "PH_PY118":  Set oTempClass = New PH_PY118: Call oTempClass.LoadForm      '//급상여Email발송
        Case "PH_PY119":  Set oTempClass = New PH_PY119: Call oTempClass.LoadForm      '//급상여Email발송
        Case "PH_PY120":  Set oTempClass = New PH_PY120: Call oTempClass.LoadForm      '//급상여소급분 생성
        Case "PH_PY121":  Set oTempClass = New PH_PY121: Call oTempClass.LoadForm      '//평가가급액 등록
        Case "PH_PY122":  Set oTempClass = New PH_PY122: Call oTempClass.LoadForm      '//급상여출력 개인부서설정등록
        Case "PH_PY123":  Set oTempClass = New PH_PY123: Call oTempClass.LoadForm      '//가압류등록
        Case "PH_PY125":  Set oTempClass = New PH_PY125: Call oTempClass.LoadForm      '//개인별 퇴직연금 설정
        Case "PH_PY127":  Set oTempClass = New PH_PY127: Call oTempClass.LoadForm      '//개인별 4대보험 보수월액 및 정산금액입력
        Case "PH_PY130":  Set oTempClass = New PH_PY130: Call oTempClass.LoadForm      '//팀별 성과급차등 등급등록
        Case "PH_PY131":  Set oTempClass = New PH_PY131: Call oTempClass.LoadForm      '//성과급차등 계수등록
        Case "PH_PY132":  Set oTempClass = New PH_PY132: Call oTempClass.LoadForm      '//성과급차등 개인별 계산
        Case "PH_PY133":  Set oTempClass = New PH_PY133: Call oTempClass.LoadForm      '//연봉제 횟차관리
        Case "PH_PY134":  Set oTempClass = New PH_PY134: Call oTempClass.LoadForm      '//소득세/주민세 조정관리
        Case "PH_PY129":  Set oTempClass = New PH_PY129: Call oTempClass.LoadForm      '//개인퇴직연금(DC형) 계산
        
        '---------------------------------------------------------------------------------------
        '// 급여 - 리포트
        '---------------------------------------------------------------------------------------
        Case "PH_PY625":  Set oTempClass = New PH_PY625: Call oTempClass.LoadForm      '//세탁자명부
        Case "PH_PY630":  Set oTempClass = New PH_PY630: Call oTempClass.LoadForm      '//사원별노조비공제현황
        Case "PH_PY700":  Set oTempClass = New PH_PY700: Call oTempClass.LoadForm      '//급여지급대장
        Case "PH_PY710":  Set oTempClass = New PH_PY710: Call oTempClass.LoadForm      '//상여지급대장장
        Case "PH_PY715":  Set oTempClass = New PH_PY715: Call oTempClass.LoadForm      '//급여부서별집계대장
        Case "PH_PY720":  Set oTempClass = New PH_PY720: Call oTempClass.LoadForm      '//상여부서별집계대장
        Case "PH_PY725":  Set oTempClass = New PH_PY725: Call oTempClass.LoadForm      '//급여직급별집계대장
        Case "PH_PY740":  Set oTempClass = New PH_PY740: Call oTempClass.LoadForm      '//상여직급별집계대장
        Case "PH_PY730":  Set oTempClass = New PH_PY730: Call oTempClass.LoadForm      '//급여봉투출력
        Case "PH_PY735":  Set oTempClass = New PH_PY735: Call oTempClass.LoadForm      '//상여봉투출력
        Case "PH_PY745":  Set oTempClass = New PH_PY745: Call oTempClass.LoadForm      '//연간지급현황
        Case "PH_PY750":  Set oTempClass = New PH_PY750: Call oTempClass.LoadForm      '//근로소득징수현황
        Case "PH_PY755":  Set oTempClass = New PH_PY755: Call oTempClass.LoadForm      '//동호회가입현황
        Case "PH_PY760":  Set oTempClass = New PH_PY760: Call oTempClass.LoadForm      '//평균임금및퇴직금산출내역서
        Case "PH_PY765":  Set oTempClass = New PH_PY765: Call oTempClass.LoadForm      '//급여증감내역서
        Case "PH_PY770":  Set oTempClass = New PH_PY770: Call oTempClass.LoadForm      '//퇴직소득원천징수영수증출력
        Case "PH_PY775":  Set oTempClass = New PH_PY775: Call oTempClass.LoadForm      '//개인별년차현황
        Case "PH_PY776":  Set oTempClass = New PH_PY776: Call oTempClass.LoadForm      '//잔여년차현황(퇴직용)
        Case "PH_PY780":  Set oTempClass = New PH_PY780: Call oTempClass.LoadForm      '//월고용보험내역
        Case "PH_PY785":  Set oTempClass = New PH_PY785: Call oTempClass.LoadForm      '//월국민연금내역
        Case "PH_PY790":  Set oTempClass = New PH_PY790: Call oTempClass.LoadForm      '//월건강보험내역
        Case "PH_PY795":  Set oTempClass = New PH_PY795: Call oTempClass.LoadForm      '//연간부서별급여내역
        Case "PH_PY800":  Set oTempClass = New PH_PY800: Call oTempClass.LoadForm      '//인건비지급자료
        Case "PH_PY805":  Set oTempClass = New PH_PY805: Call oTempClass.LoadForm      '//급여수당변동내역
        Case "PH_PY810":  Set oTempClass = New PH_PY810: Call oTempClass.LoadForm      '//직급별통상임금내역
        Case "PH_PY815":  Set oTempClass = New PH_PY815: Call oTempClass.LoadForm      '//평균임금내역
        Case "PH_PY820":  Set oTempClass = New PH_PY820: Call oTempClass.LoadForm      '//통상임금내역
        Case "PH_PY825":  Set oTempClass = New PH_PY825: Call oTempClass.LoadForm      '//전문직O/T현황
        Case "PH_PY830":  Set oTempClass = New PH_PY830: Call oTempClass.LoadForm      '//부서별인건비현황 (기획)
        Case "PH_PY835":  Set oTempClass = New PH_PY835: Call oTempClass.LoadForm      '//직급별O/T및수당현황
        Case "PH_PY840":  Set oTempClass = New PH_PY840: Call oTempClass.LoadForm      '//풍산전자공시자료
        Case "PH_PY845":  Set oTempClass = New PH_PY845: Call oTempClass.LoadForm      '//기간별급여지급내역
        Case "PH_PY850":  Set oTempClass = New PH_PY850: Call oTempClass.LoadForm      '//소급분지급명세서
        Case "PH_PY855":  Set oTempClass = New PH_PY855: Call oTempClass.LoadForm      '//개인별임금지급대장
        Case "PH_PY865":  Set oTempClass = New PH_PY865: Call oTempClass.LoadForm      '//고용보험현황(계산용)
        Case "PH_PY870":  Set oTempClass = New PH_PY870: Call oTempClass.LoadForm      '//담당별월O/T및수당현황
        Case "PH_PY875":  Set oTempClass = New PH_PY875: Call oTempClass.LoadForm      '//직급별수당집계대장
        Case "PH_PY716":  Set oTempClass = New PH_PY716: Call oTempClass.LoadForm      '//기간별급여부서별집계대장
        Case "PH_PY721":  Set oTempClass = New PH_PY721: Call oTempClass.LoadForm      '//기간별상여부서별집계대장
        Case "PH_PY717":  Set oTempClass = New PH_PY717: Call oTempClass.LoadForm      '//기간별급여반별집계대장
        Case "PH_PY718":  Set oTempClass = New PH_PY718: Call oTempClass.LoadForm      '생산완료금액대비O/T현황(2014.08.19 송명규)
        Case "PH_PY701":  Set oTempClass = New PH_PY701: Call oTempClass.LoadForm      '//급여지급대장(노조용)
        
        Case "PH_PYA10":  Set oTempClass = New PH_PYA10: Call oTempClass.LoadForm      '//급여지급대장(부서)
        Case "PH_PYA20":  Set oTempClass = New PH_PYA20: Call oTempClass.LoadForm      '//급여부서별집계대장(부서)
        Case "PH_PYA30":  Set oTempClass = New PH_PYA30: Call oTempClass.LoadForm      '//상여지급대장(부서)
        Case "PH_PYA40":  Set oTempClass = New PH_PYA40: Call oTempClass.LoadForm      '//상여부서별집계대장(부서)
        Case "PH_PYA50":  Set oTempClass = New PH_PYA50: Call oTempClass.LoadForm      '//DC전환자부담금지급내역
        Case "PH_PYA75":  Set oTempClass = New PH_PYA75: Call oTempClass.LoadForm      '//교통비외수당지급대장
        
        '---------------------------------------------------------------------------------------
        '// 정산
        '---------------------------------------------------------------------------------------
        Case "PH_PY401":  Set oTempClass = New PH_PY401: Call oTempClass.LoadForm      '전근무지등록
        Case "PH_PY402":  Set oTempClass = New PH_PY402: Call oTempClass.LoadForm      '정산기초자료 등록
        Case "PH_PY405":  Set oTempClass = New PH_PY405: Call oTempClass.LoadForm      '의료비등록
        Case "PH_PY407":  Set oTempClass = New PH_PY407: Call oTempClass.LoadForm      '기부금등록
        Case "PH_PY409":  Set oTempClass = New PH_PY409: Call oTempClass.LoadForm      '기부금조정명세등록
        Case "PH_PY411":  Set oTempClass = New PH_PY411: Call oTempClass.LoadForm      '연금.저축등소득공제등록
        Case "PH_PY413":  Set oTempClass = New PH_PY413: Call oTempClass.LoadForm      '월세액.주택임차차입금자료 등록
        Case "PH_PY415":  Set oTempClass = New PH_PY415: Call oTempClass.LoadForm      '정산계산
        Case "PH_PY417":  Set oTempClass = New PH_PY417: Call oTempClass.LoadForm      '정산 은행파일생성
        Case "PH_PY980":  Set oTempClass = New PH_PY980: Call oTempClass.LoadForm      '신고_근로소득지급명세서자료작성
        Case "PH_PY985":  Set oTempClass = New PH_PY985: Call oTempClass.LoadForm      '신고_의료비지급명세서자료작성
        Case "PH_PY990":  Set oTempClass = New PH_PY990: Call oTempClass.LoadForm      '신고_기부금명세서자료작성
        Case "PH_PY995":  Set oTempClass = New PH_PY995: Call oTempClass.LoadForm      '신고_퇴직소득지급명세서자료작성
        Case "PH_PY419":  Set oTempClass = New PH_PY419: Call oTempClass.LoadForm      '표준세액적용대상자등록
        
        Case "PH_PY910":  Set oTempClass = New PH_PY910: Call oTempClass.LoadForm      '소득공제신고서출력
        Case "PH_PY915":  Set oTempClass = New PH_PY915: Call oTempClass.LoadForm      '근로소득원천징수부출력
        Case "PH_PY920":  Set oTempClass = New PH_PY920: Call oTempClass.LoadForm      '원천징수영수증출력
        Case "PH_PY925":  Set oTempClass = New PH_PY925: Call oTempClass.LoadForm      '기부금명세서출력
        Case "PH_PY930":  Set oTempClass = New PH_PY930: Call oTempClass.LoadForm      '전산징수및환급대장
        Case "PH_PY931":  Set oTempClass = New PH_PY931: Call oTempClass.LoadForm      '표준세액적용대상자조회
        Case "PH_PY932":  Set oTempClass = New PH_PY932: Call oTempClass.LoadForm      '전근무지등록현황
        Case "PH_PY933":  Set oTempClass = New PH_PY933: Call oTempClass.LoadForm      '보수총액신고기초자료
        Case "PH_PYA55":  Set oTempClass = New PH_PYA55: Call oTempClass.LoadForm      '정산징수및환급대장(집계)
        Case "PH_PYA70":  Set oTempClass = New PH_PYA70: Call oTempClass.LoadForm      '소득세원천징수세액조정신청서출력
        
        Case "ZPY341":  Set oTempClass = New ZPY341: Call oTempClass.LoadForm       '월별자료생성
        Case "ZPY343":  Set oTempClass = New ZPY343: Call oTempClass.LoadForm       '사원별월자료집계
        
        Case "ZPY421":  Set oTempClass = New ZPY421: Call oTempClass.LoadForm       '퇴직소득전산매체수록
        Case "ZPY501":  Set oTempClass = New ZPY501: Call oTempClass.LoadForm       '소득공제항목 등록
        Case "ZPY502":  Set oTempClass = New ZPY502: Call oTempClass.LoadForm       '종전근무지등록
        Case "ZPY503":  Set oTempClass = New ZPY503: Call oTempClass.LoadForm       '정산세액 계산
        Case "ZPY504":  Set oTempClass = New ZPY504: Call oTempClass.LoadForm       '정산결과 조회
        Case "ZPY505":  Set oTempClass = New ZPY505: Call oTempClass.LoadForm       '기부금명세등록
        Case "ZPY506":  Set oTempClass = New ZPY506: Call oTempClass.LoadForm       '의료비명세등록
        Case "ZPY507":  Set oTempClass = New ZPY507: Call oTempClass.LoadForm       '정산결과 조회(전체)
        Case "ZPY508":  Set oTempClass = New ZPY508: Call oTempClass.LoadForm       '연금.저축 명세 등록
        Case "ZPY509":  Set oTempClass = New ZPY509: Call oTempClass.LoadForm       '정산자료마감
        Case "ZPY510":  Set oTempClass = New ZPY510: Call oTempClass.LoadForm       '종전근무지 일괄생성
        Case "ZPY521":  Set oTempClass = New ZPY521: Call oTempClass.LoadForm       '근로 소득 전산 매체
        Case "ZPY522":  Set oTempClass = New ZPY522: Call oTempClass.LoadForm       '의료/기부금전산매체
        
        Case "RPY401":  Set oTempClass = New RPY401: Call oTempClass.LoadForm       '퇴직원천징수영수증
        Case "RPY501":  Set oTempClass = New RPY501: Call oTempClass.LoadForm       '월별자료현황
        Case "RPY502":  Set oTempClass = New RPY502: Call oTempClass.LoadForm       '종전근무지현황
        Case "RPY503":  Set oTempClass = New RPY503: Call oTempClass.LoadForm       '근로소득 원천징수부
        Case "RPY504":  Set oTempClass = New RPY504: Call oTempClass.LoadForm       '근로소득 원천영수증
        Case "RPY505":  Set oTempClass = New RPY505: Call oTempClass.LoadForm       '소득자료집계표
        Case "RPY506":  Set oTempClass = New RPY506: Call oTempClass.LoadForm       '정산징수환급대장
        Case "RPY508":  Set oTempClass = New RPY508: Call oTempClass.LoadForm       '연말정산집계표
        Case "RPY509":  Set oTempClass = New RPY509: Call oTempClass.LoadForm       '갑근세신고검토표
        Case "RPY510":  Set oTempClass = New RPY510: Call oTempClass.LoadForm       '비과세근로소득명세서
        Case "RPY511":  Set oTempClass = New RPY511: Call oTempClass.LoadForm       '기부금 명세서
        
        '---------------------------------------------------------------------------------------
        '// 기타
        '---------------------------------------------------------------------------------------
        Case "PH_PY301":  Set oTempClass = New PH_PY301: Call oTempClass.LoadForm       '//학자금신청등록
        Case "PH_PY302":  Set oTempClass = New PH_PY302: Call oTempClass.LoadForm       '//학자금지급완료처리
        Case "PH_PY303":  Set oTempClass = New PH_PY303: Call oTempClass.LoadForm       '//학자금은행파일생성
        Case "PH_PY305":  Set oTempClass = New PH_PY305: Call oTempClass.LoadForm       '//학자금신청서
        Case "PH_PY306":  Set oTempClass = New PH_PY306: Call oTempClass.LoadForm       '//학자금신청내역(개인별)
        Case "PH_PY307":  Set oTempClass = New PH_PY307: Call oTempClass.LoadForm       '//학자금신청내역(분기별)
        Case "PH_PY309":  Set oTempClass = New PH_PY309: Call oTempClass.LoadForm       '//대부금등록
        Case "PH_PY310":  Set oTempClass = New PH_PY310: Call oTempClass.LoadForm       '//대부금개별상환
        Case "PH_PY311":  Set oTempClass = New PH_PY311: Call oTempClass.LoadForm       '//통근버스운행등록
        Case "PH_PY312":  Set oTempClass = New PH_PY312: Call oTempClass.LoadForm       '//버스요금 등록(창원)
        Case "PH_PY313":  Set oTempClass = New PH_PY313: Call oTempClass.LoadForm       '//대부금계산
        Case "PH_PY314":  Set oTempClass = New PH_PY314: Call oTempClass.LoadForm       '//대부금계산내역조회
        Case "PH_PY030":  Set oTempClass = New PH_PY030: Call oTempClass.LoadForm       '//공용등록
        Case "PH_PY031":  Set oTempClass = New PH_PY031: Call oTempClass.LoadForm       '//출장등록
        Case "PH_PY032":  Set oTempClass = New PH_PY032: Call oTempClass.LoadForm       '//사용외출등록
        Case "PH_PY315":  Set oTempClass = New PH_PY315: Call oTempClass.LoadForm       '//개인별대부금잔액현황
        Case "PH_PY034":  Set oTempClass = New PH_PY034: Call oTempClass.LoadForm       '//공용분개처리(2013.10.14 송명규 추가)
        Case "PH_PYA60":  Set oTempClass = New PH_PYA60: Call oTempClass.LoadForm       '//학자금신청내역(집계)
        End Select
        

    End If

    Exit Sub
    
CreateUSERForm_Error:
    Sbo_Application.SetStatusBarMessage "CreateUSERForm_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

'//시스템폼 아이템 이벤트
Private Sub CreateSYSTEMForm(ByRef pval As SAPbouiCOM.ItemEvent)

    Dim oTempClass          As Variant
    
    On Error GoTo CreateSYSTEMForm_Error
    
    If (pval.BeforeAction = True) Then
        If (pval.EventType = et_FORM_LOAD) Then
            Select Case pval.FormTypeEx
             
'            Case "-60100"       '//인사관리>사원마스터데이터 (사용자 정의 필드)
'                Set oTempClass = New SM60100: oTempClass.LoadForm (pval.FormUID): AddForms oTempClass, pval.FormUID, pval.FormTypeEx

            End Select

        End If
    End If
    Exit Sub
    
CreateSYSTEMForm_Error:
    Sbo_Application.SetStatusBarMessage "CreateSYSTEMForm_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub DoSomething()
    
    On Error GoTo DoSomething_Error
    
    If Sbo_Application.Menus("1024").SubMenus("1026").Enabled = True Then
        Call Sbo_Application.Menus("1024").SubMenus("1026").Activate
    End If
    Exit Sub
    
DoSomething_Error:
    Sbo_Application.SetStatusBarMessage "DoSomething_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub





























