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
'//  SAP MANAGE UI API 6.7 SDK Sample
'//****************************************************************************
'//  File          :      ZZMDC.cls
'//  Create Date   : 2007.05.14
'//  Modified Date : 2010.08.23
'//  Creator       : Hibi
'//****************************************************************************
Option Explicit
Option Base 1
Private WithEvents Sbo_Application      As SAPbouiCOM.Application   'SBO application
Attribute Sbo_Application.VB_VarHelpID = -1

Private Sub Class_Initialize()
    MDC_Globals.ProgramType = "Local"
'    MDC_Globals.ProgramType = "Server"
    Call Initialize '//초기화
    Call SetApplication '//Sbo_Application초기화
    Call ConnectToSBO '//Sbo_Company초기화
    Call SetPath '//경로설정
    Call ConnectShareFolder '//공유폴더연결
    Call AddMenuXml '//메뉴추가
    Call DoSomething
    Sbo_Application.StatusBar.SetText "풍산홀딩스 애드온이 초기화되었습니다.", bmt_Short, smt_Success
End Sub

Private Sub Initialize()
    SubMain.FormCurrentCount = 0
    SubMain.FormTotalCount = 0
    Set SubMain.ClassList = New Collection
    SubMain.ShareFolderPath = ""
    SubMain.ServerPath = ""
End Sub

Private Sub SetApplication()
    Dim SboGuiApi As New SAPbouiCOM.SboGuiApi
    Dim ConnectionString As String
    ConnectionString = Command
    If Trim(ConnectionString) = "" Then
        ConnectionString = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056"
    End If
    Call SboGuiApi.Connect(ConnectionString)
    Set Sbo_Application = SboGuiApi.GetApplication()
    Set SubMain.Sbo_Application = Sbo_Application
    Sbo_Application.SetStatusBarMessage "Addon을 시작합니다.", bmt_Short, False
End Sub

Public Sub ConnectToSBO()
On Error GoTo ConnectToSBO_Error

    Set Sbo_Company = Sbo_Application.Company.GetDICompany()
    
    If Sbo_Company.Connected = False Then
        Sbo_Application.MessageBox "DB 접속실패"
        End
    End If
    
    Sbo_Application.SetStatusBarMessage "DB 접속성공", bmt_Short, False
    
    Exit Sub

ConnectToSBO_Error:
    Call MsgBox("DB 접속실패", vbOKOnly, "SAP Business One")
    End
End Sub

'Private Sub ConnectToSBO()
'    If Not SetConnectionContext = 0 Then
'        Sbo_Application.MessageBox "Failed setting a connection to DI API"
'        End 'Add-On Application 종료
'    End If
'    If Not ConnectToSbo_Company = 0 Then
'        Sbo_Application.MessageBox "Failed connecting to the company's Database"
'        End 'Add-On Application 종료
'    End If
'End Sub
'
'Private Function SetConnectionContext() As Long
'    Dim sCookie As String
'    Dim sConnectionContext As String
'    Dim lRetCode As Long
'    Set Sbo_Company = New SAPbobsCOM.Company
'    Sbo_Company.AddonIdentifier = "5645523035496D706C656D656E746174696F6E3A5430333834373037313836A1E9C35201042394B16C07851EA9BB0BA0408B6D"
'    sCookie = Sbo_Company.GetContextCookie
'    sConnectionContext = Sbo_Application.Company.GetConnectionContext(sCookie)
'    If Sbo_Company.Connected = True Then
'        Sbo_Application.MessageBox Sbo_Company.CompanyDB & " now disconnect !!"
'        Sbo_Company.Disconnect
'    End If
'    SetConnectionContext = Sbo_Company.SetSboLoginContext(sConnectionContext)
'End Function
'
'Private Function ConnectToSbo_Company() As Long
'    ConnectToSbo_Company = Sbo_Company.Connect
'End Function

Public Sub SetPath()
    If (ProgramType = "Local") Then '//로컬일경우
        ServerPath = App.Path
    ElseIf (ProgramType = "Server") Then '//서버일경우
        ServerPath = MDC_PS_Common.GetValue("EXEC Profile_SELECT 'SERVERINFO'", 0, 1)
    End If
End Sub

Private Sub ConnectShareFolder()
    If (ProgramType = "Local") Then
        SubMain.ShareFolderPath = ServerPath & "\PathINI\"
    ElseIf (ProgramType = "Server") Then
        '//공유폴더연결
        Dim Shell As Object
        Set Shell = CreateObject("Wscript.Shell")
        Call Shell.Run("NET USE * /D /YES", 1, True) '모든 네트워크 드라이브 연결 끊기(2011.12.14 송명규 추가)
'        Call Shell.Run("NET USE " & MDC_PS_Common.GetValue("EXEC Profile_SELECT 'SERVERINFO'", 0, 1) & "\IPC$" & " /D", 1, True)
'        Call Shell.Run("NET USE " & MDC_PS_Common.GetValue("EXEC Profile_SELECT 'SERVERINFO'", 0, 1) & MDC_PS_Common.GetValue("EXEC Profile_SELECT 'SERVERINFO'", 1, 1) & " /D", 1, True)
        Call SetDrive(ServerPath, MDC_PS_Common.GetValue("EXEC Profile_SELECT 'SERVERINFO'", 1, 1), MDC_PS_Common.GetValue("EXEC Profile_SELECT 'SERVERINFO'", 2, 1), MDC_PS_Common.GetValue("EXEC Profile_SELECT 'SERVERINFO'", 3, 1), MDC_PS_Common.GetValue("EXEC Profile_SELECT 'SERVERINFO'", 4, 1)) '//DB에서 참조방식으로변경
        SubMain.ShareFolderPath = MDC_PS_Common.GetValue("EXEC Profile_SELECT 'SERVERINFO'", 5, 1)
    End If
End Sub

Private Sub AddMenuXml()
    Dim FileName As String
    
    If Sbo_Application.Company.UserName = "PSH06" Or Sbo_Application.Company.UserName = "PSH07" Then '특정 계정은 메인메뉴 최소화
        FileName = "Menu_KOR2.xml"
    ElseIf Sbo_Application.Company.UserName = "71090" Or Sbo_Application.Company.UserName = "66302" Then
        FileName = "Menu_KOR3.xml"
    Else
        FileName = "Menu_KOR.xml"
    End If
    
    Dim oXmlDoc As New MSXML2.DOMDocument
    oXmlDoc.Load (ShareFolderPath & FileName)
    Sbo_Application.LoadBatchActions oXmlDoc.xml
End Sub

Private Sub TerminateApplication()
    Dim ws
    Dim oDrives
    Dim i As Long
    Dim oTempClass As Variant
    
    '//네트워크 드라이버 해제
    Set ws = CreateObject("WScript.Network")
    Set oDrives = ws.EnumNetworkDrives
    For i = 0 To oDrives.Count - 1
        If oDrives(i) = MDC_PS_Common.GetValue("EXEC Profile_SELECT 'SERVERINFO'", 4, 1) Then
            ws.RemoveNetworkDrive oDrives.Item(i)
        End If
'        If oDrives(i) = "W:" Then
'            ws.RemoveNetworkDrive oDrives.Item(i)
'        End If
    Next i
    
    '//오브젝트제거
    If ClassList.Count > 0 Then
        For i = 0 To ClassList.Count
            Set oTempClass = ClassList.Item(i)
            ClassList.Remove i
            Set oTempClass = Nothing
        Next i
    End If
    Sbo_Company.Disconnect
    Set Sbo_Company = Nothing
    Set g_ERPDMS = Nothing
    End
End Sub

Private Sub Sbo_Application_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes)
    Select Case EventType
        Case aet_ShutDown:
            Call TerminateApplication
        Case aet_CompanyChanged:
            Call TerminateApplication
        Case aet_LanguageChanged:
            'Call TerminateApplication
        Case aet_ServerTerminition:
            Call TerminateApplication
    End Select
End Sub

Private Sub Sbo_Application_ItemEvent(ByVal FormUID As String, pval As SAPbouiCOM.ItemEvent, BubbleEvent As Boolean)
On Error GoTo Sbo_Application_ItemEvent_Error
    Dim oTempClass          As Variant
    Call CreateSystemForm(pval) '//시스템폼생성
    If (pval.FormTypeEx = "60012" Or pval.FormTypeEx = "FormEditorType") Then
        Exit Sub
    End If
    If Left(pval.FormUID, 2) = "PS" Then '//풍산애드온에의해 생성된폼
        Set oTempClass = ClassList.Item(FormUID)
        If (oTempClass.oForm01 Is Nothing) Then
            Exit Sub
        Else
            Call oTempClass.Raise_ItemEvent(FormUID, pval, BubbleEvent)
        End If
    ElseIf Left(pval.FormUID, 2) = "F_" Then '//시스템폼
        If ValidateForm(pval.FormTypeEx) Then
            Set oTempClass = ClassList.Item(FormUID)
            Call oTempClass.Raise_ItemEvent(FormUID, pval, BubbleEvent)
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
On Error GoTo Sbo_Application_MenuEvent_Error
    Dim oTempClass          As Variant
    Call CreateUserForm(pval) '//사용자폼생성
    
    Dim FormUID As String
    FormUID = Sbo_Application.Forms.ActiveForm.uniqueID
    
    If Left(FormUID, 2) = "PS" Then '//풍산애드온에의해 생성된폼
        Set oTempClass = ClassList.Item(FormUID)
        If (oTempClass.oForm01 Is Nothing) Then
            Exit Sub
        Else
            Call oTempClass.Raise_MenuEvent(FormUID, pval, BubbleEvent)
        End If
    ElseIf Left(FormUID, 2) = "F_" Then '//시스템폼
        If ValidateForm(Sbo_Application.Forms.ActiveForm.TypeEx) Then
            Set oTempClass = ClassList.Item(FormUID)
            Call oTempClass.Raise_MenuEvent(FormUID, pval, BubbleEvent)
        End If
    End If
    Exit Sub
Sbo_Application_MenuEvent_Error:
    If Err.Number = 5 Then
        '//모두닫기메뉴실행시 이전에존재하는 PS폼에의해발생
    Else
        Sbo_Application.SetStatusBarMessage "Sbo_Application_MenuEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    End If
End Sub

Private Sub Sbo_Application_FormDataEvent(BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, BubbleEvent As Boolean)
On Error GoTo Sbo_Application_FormDataEvent_Error
    Dim oTempClass As Variant
    Dim FormUID As String
    FormUID = BusinessObjectInfo.FormUID
    
    If Left(FormUID, 2) = "PS" Then '//풍산애드온에의해 생성된폼
        Set oTempClass = ClassList.Item(FormUID)
        If (oTempClass.oForm01 Is Nothing) Then
            Exit Sub
        Else
            Call oTempClass.Raise_FormDataEvent(FormUID, BusinessObjectInfo, BubbleEvent)
        End If
    ElseIf Left(FormUID, 2) = "F_" Then '//시스템폼
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
On Error GoTo Sbo_Application_RightClickEvent_Error
    Dim oTempClass As Variant
    Dim FormUID As String
    FormUID = eventInfo.FormUID
    If Left(FormUID, 2) = "PS" Then '//풍산애드온에의해 생성된폼
        Set oTempClass = ClassList.Item(FormUID)
        If (oTempClass.oForm01 Is Nothing) Then
            Exit Sub
        Else
            Call oTempClass.Raise_RightClickEvent(FormUID, eventInfo, BubbleEvent)
        End If
    ElseIf Left(FormUID, 2) = "F_" Then '//시스템폼
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
On Error GoTo ValidateForm_Error:
    ValidateForm = True
    '//판매관리
    If FormType = "149" Then
    ElseIf FormType = "139" Then
    ElseIf FormType = "140" Then
    ElseIf FormType = "180" Then
    ElseIf FormType = "133" Then
    ElseIf FormType = "179" Then
    ElseIf FormType = "60091" Then
    '//구매관리
    ElseIf FormType = "142" Then
    ElseIf FormType = "143" Then
    ElseIf FormType = "182" Then
    ElseIf FormType = "141" Then
    ElseIf FormType = "181" Then
    ElseIf FormType = "60092" Then
    '//회계관리
    ElseIf FormType = "804" Then
    ElseIf FormType = "392" Then
    ElseIf FormType = "393" Then
    '//자금관리
    ElseIf FormType = "146" Then
    ElseIf FormType = "196" Then
    ElseIf FormType = "170" Then
    ElseIf FormType = "426" Then
    ElseIf FormType = "60051" Then
    '//기타
    ElseIf FormType = "41" Then
    ElseIf FormType = "42" Then
    ElseIf FormType = "720" Then
    ElseIf FormType = "721" Then
    ElseIf FormType = "150" Then '품목마스터
    Else
        ValidateForm = False
    End If
    Exit Function
ValidateForm_Error:
    Sbo_Application.SetStatusBarMessage "ValidateForm_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub CreateUserForm(ByRef pval As SAPbouiCOM.IMenuEvent)
On Error GoTo CreateUserForm_Error
    Dim oTempClass          As Variant
    
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If (pval.BeforeAction = True) Then
    
        sQry = " EXEC Z_PS_FormCount '" & MDC_PS_Common.User_MSTCOD & "','" & pval.MenuUID & "'"
        Call oRecordSet01.DoQuery(sQry)
        
        '//강상규
        If (pval.MenuUID = "PS_SM010") Then '//재고조회
            Set oTempClass = New PS_SM010
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SM020") Then '//재고조회
            Set oTempClass = New PS_SM020
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD030") Then '//출하요청
            Set oTempClass = New PS_SD030
            Call oTempClass.LoadForm("", "출하요청")
        ElseIf (pval.MenuUID = "PS_SD031") Then '//선출요청
            Set oTempClass = New PS_SD030
            Call oTempClass.LoadForm("", "선출요청")
        ElseIf (pval.MenuUID = "PS_SD040") Then '//납품
            Set oTempClass = New PS_SD040
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM065") Then '//고객별휘팅할인율등록
            Set oTempClass = New PS_MM065
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM090") Then '//자재기타출고                               '//성수추가
            Set oTempClass = New PS_MM090
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP002") Then '//불량중분류코드등록
            Set oTempClass = New PS_PP002
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP003") Then '//불량상세분류코드등록
            Set oTempClass = New PS_PP003
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP004") Then '//품목별표준공정등록
            Set oTempClass = New PS_PP004
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP030") Then '//작업지시등록
            Set oTempClass = New PS_PP030
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP040") Then '//작업일보등록(작지)
            Set oTempClass = New PS_PP040
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP041") Then '//작업일보등록(공정)
            Set oTempClass = New PS_PP041
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP070") Then '//휘팅벌크포장
            Set oTempClass = New PS_PP070
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP080") Then '//생산완료등록(작지)
            Set oTempClass = New PS_PP080
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP090") Then '//멀티패킹등록                                 '//성수추가
            Set oTempClass = New PS_PP090
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP120") Then '//PSMT임가공내역등록 및 비용전기         '//성수추가
            Set oTempClass = New PS_PP120
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM180") Then '//입고(배치번호관리품목)
            Set oTempClass = New PS_MM180
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD380") Then '//레포트-납품명세서
            Set oTempClass = New PS_SD380
            Call oTempClass.LoadForm
        End If
        
        '//RYJ
        '//MM - Object
        If (pval.MenuUID = "PS_MM005") Then
            Set oTempClass = New PS_MM005
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM005_1") Then
            Set oTempClass = New PS_MM005
            Call oTempClass.LoadForm("PS_MM005_1")
        ElseIf (pval.MenuUID = "PS_MM010") Then
            Set oTempClass = New PS_MM010
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM030") Then
            Set oTempClass = New PS_MM030
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM040") Then
            Set oTempClass = New PS_MM040
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM050") Then
            Set oTempClass = New PS_MM050
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM070") Then
            Set oTempClass = New PS_MM070
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM095") Then
            Set oTempClass = New PS_MM095
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM140") Then
            Set oTempClass = New PS_MM140
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM145") Then
            Set oTempClass = New PS_MM145
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM150") Then
            Set oTempClass = New PS_MM150
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM150_1") Then
            Set oTempClass = New PS_MM150
            Call oTempClass.LoadForm("PS_MM150_1")
        ElseIf (pval.MenuUID = "PS_MM152") Then
            Set oTempClass = New PS_MM152
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM155") Then
            Set oTempClass = New PS_MM155
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM155_1") Then
            Set oTempClass = New PS_MM155
            Call oTempClass.LoadForm("PS_MM155_1")
        ElseIf (pval.MenuUID = "PS_MM160") Then
            Set oTempClass = New PS_MM160
            Call oTempClass.LoadForm
        '//SD - Object
        ElseIf (pval.MenuUID = "PS_SD010") Then
            Set oTempClass = New PS_SD010
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD080") Then
            Set oTempClass = New PS_SD080
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD082") Then
            Set oTempClass = New PS_SD082
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD920") Then
            Set oTempClass = New PS_SD920
            Call oTempClass.LoadForm
        '//PP - Object
        ElseIf (pval.MenuUID = "PS_PP010") Then
            Set oTempClass = New PS_PP010
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP020") Then
            Set oTempClass = New PS_PP020
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP045") Then
            Set oTempClass = New PS_PP045
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP091") Then
            Set oTempClass = New PS_PP091
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP918") Then
            Set oTempClass = New PS_PP918
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP920") Then
            Set oTempClass = New PS_PP920
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP950") Then
            Set oTempClass = New PS_PP950
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP951") Then
            Set oTempClass = New PS_PP951
            Call oTempClass.LoadForm
        '//CO - Object
        ElseIf (pval.MenuUID = "PS_CO110") Then
            Set oTempClass = New PS_CO110
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_CO120") Then
            Set oTempClass = New PS_CO120
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_CO130") Then
            Set oTempClass = New PS_CO130
            Call oTempClass.LoadForm
        '//FI - Object
        ElseIf (pval.MenuUID = "PS_FI420") Then
            Set oTempClass = New PS_FI420
            Call oTempClass.LoadForm
        '//MM - Report
        ElseIf (pval.MenuUID = "PS_MM212") Then
            Set oTempClass = New PS_MM212
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM224") Then
            Set oTempClass = New PS_MM224
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM237") Then
            Set oTempClass = New PS_MM237
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM245") Then
            Set oTempClass = New PS_MM245
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM928") Then
            Set oTempClass = New PS_MM928
            Call oTempClass.LoadForm
        '//FI - Report
        ElseIf (pval.MenuUID = "PS_FI220") Then
            Set oTempClass = New PS_FI220
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_FI900") Then
            Set oTempClass = New PS_FI900
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_FI901") Then
            Set oTempClass = New PS_FI901
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_FI902") Then
            Set oTempClass = New PS_FI902
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_FI929") Then
            Set oTempClass = New PS_FI929
            Call oTempClass.LoadForm
        '//PP - Report
        ElseIf (pval.MenuUID = "PS_PP923") Then
            Set oTempClass = New PS_PP923
            Call oTempClass.LoadForm
        End If
        
        '// 윤제형
        If (pval.MenuUID = "PS_PP071") Then     '//작지조회
            Set oTempClass = New PS_PP071
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP075") Then     '//휘팅이동등록
            Set oTempClass = New PS_PP075
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP077") Then '//휘팅포장등록
            Set oTempClass = New PS_PP077
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP078") Then '//휘팅포장등록취소
            Set oTempClass = New PS_PP078
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP777") Then '//벌크반품등록
            Set oTempClass = New PS_PP777
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP778") Then '//벌크반품이동등록
            Set oTempClass = New PS_PP778
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP913") Then '//장비가동시간등록
            Set oTempClass = New PS_PP913
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM010") Then '//검사사양서등록
            Set oTempClass = New PS_QM010
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM020") Then '//검사성적서등록
            Set oTempClass = New PS_QM020
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM030") Then '//원소재성분입력
            Set oTempClass = New PS_QM030
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM040") Then '//검사성적서출력
            Set oTempClass = New PS_QM040
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM902") Then '//실패비용 산출단가등록
            Set oTempClass = New PS_QM902
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM203") Then '//입고 현황
            Set oTempClass = New PS_MM203
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM204") Then '//검수 현황
            Set oTempClass = New PS_MM204
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM207") Then '//출고 대장
            Set oTempClass = New PS_MM207
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM209") Then '//수불 명세서
            Set oTempClass = New PS_MM209
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM214") Then '//품목수불 명세서
            Set oTempClass = New PS_MM214
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM210") Then '//외주스크랩 입고현황
            Set oTempClass = New PS_MM210
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM211") Then '//수불 대장
            Set oTempClass = New PS_MM211
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM249") Then '//제품수불부
            Set oTempClass = New PS_MM249
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM913") Then '//수불현황(중분류별 집계)
            Set oTempClass = New PS_MM913
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM917") Then '//원재료(부품)입고현황
            Set oTempClass = New PS_MM917
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM915") Then '//SCRAP 수불부
            Set oTempClass = New PS_MM915
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_FI030") Then '//어음발행리스트
            Set oTempClass = New PS_FI030
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_FI040") Then '//만기일자별 받을어음 현황
            Set oTempClass = New PS_FI040
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_FI050") Then '//받을어음추심의뢰서
            Set oTempClass = New PS_FI050
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_FI060") Then '//받을어음수탁(발송)대장
            Set oTempClass = New PS_FI060
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_FI070") Then '//기간별지급어음현황
            Set oTempClass = New PS_FI070
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_FI080") Then '//일계표
            Set oTempClass = New PS_FI080
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_FI090") Then '//월계표
            Set oTempClass = New PS_FI090
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_FI100") Then '//계정별보조원장
            Set oTempClass = New PS_FI100
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_FI110") Then '//보조재료비현황
            Set oTempClass = New PS_FI110
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_FI120") Then '//노무비현황
            Set oTempClass = New PS_FI120
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_FI130") Then '//제조경비현황
            Set oTempClass = New PS_FI130
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_FI140") Then '//판관비현황
            Set oTempClass = New PS_FI140
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_FI150") Then '//영업외손익,특별손익명세서
            Set oTempClass = New PS_FI150
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_FI160") Then '//외상매입금 잔액현황
            Set oTempClass = New PS_FI160
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_FI170") Then '//신용카드 사용내역
            Set oTempClass = New PS_FI170
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_FI180") Then '//보조원장
            Set oTempClass = New PS_FI180
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_FI924") Then '//거래처별 회전일현황
            Set oTempClass = New PS_FI924
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_FI910") Then '//매출채권 연령분석(받을어음)
            Set oTempClass = New PS_FI910
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_FI931") Then '//부가세집계표
            Set oTempClass = New PS_FI931
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_FI904") Then '//고정자산현황
            Set oTempClass = New PS_FI904
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_FI907") Then '//부도어음현황
            Set oTempClass = New PS_FI907
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_FI909") Then '//미수현황
            Set oTempClass = New PS_FI909
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_FI912") Then '//예산집행현황
            Set oTempClass = New PS_FI912
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_FI913") Then '//사원별접대비명세서
            Set oTempClass = New PS_FI913
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_FI918") Then '//대체출고대장
            Set oTempClass = New PS_FI918
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_FI932") Then '//차입금등록
            Set oTempClass = New PS_FI932
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_FI933") Then '//차입금대장
            Set oTempClass = New PS_FI933
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD320") Then '//품목별 판매내역서
            Set oTempClass = New PS_SD320
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD912") Then '//년간판매현황
            Set oTempClass = New PS_SD912
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD913") Then '//판매미확정대장
            Set oTempClass = New PS_SD913
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD914") Then '//판매R/M현황
            Set oTempClass = New PS_SD914
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD916") Then '//업체별 미수현황
            Set oTempClass = New PS_SD916
            Call oTempClass.LoadForm
        End If
        
        '// 김동섭
        If (pval.MenuUID = "PS_SD090") Then
            Set oTempClass = New PS_SD090
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD091") Then
            Set oTempClass = New PS_SD091
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SY001") Then
            Set oTempClass = New PS_SY001
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM170") Then
            Set oTempClass = New PS_MM170
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM060") Then
            Set oTempClass = New PS_MM060
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP001") Then
            Set oTempClass = New PS_PP001
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD001") Then
            Set oTempClass = New PS_SD001
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP035") Then
            Set oTempClass = New PS_PP035
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM130") Then
            Set oTempClass = New PS_MM130
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM132") Then
            Set oTempClass = New PS_MM132
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP006") Then
            Set oTempClass = New PS_PP006
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_CO040") Then
            Set oTempClass = New PS_CO040
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_CO060") Then
            Set oTempClass = New PS_CO060
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_CO100") Then
            Set oTempClass = New PS_CO100
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP001") Then
            Set oTempClass = New PS_PP001
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP929") Then
            Set oTempClass = New PS_PP929
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP933") Then
            Set oTempClass = New PS_PP933
            Call oTempClass.LoadForm
            
        ElseIf (pval.MenuUID = "PS_MM215") Then
            Set oTempClass = New PS_MM215
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM216") Then
            Set oTempClass = New PS_MM216
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM217") Then
            Set oTempClass = New PS_MM217
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD110") Then
            Set oTempClass = New PS_SD110
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD120") Then
            Set oTempClass = New PS_SD120
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD130") Then
            Set oTempClass = New PS_SD130
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD160") Then
            Set oTempClass = New PS_SD160
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD180") Then
            Set oTempClass = New PS_SD180
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD240") Then
            Set oTempClass = New PS_SD240
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD250") Then
            Set oTempClass = New PS_SD250
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD270") Then
            Set oTempClass = New PS_SD270
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD280") Then
            Set oTempClass = New PS_SD280
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD290") Then
            Set oTempClass = New PS_SD290
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD291") Then
            Set oTempClass = New PS_SD291
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD310") Then
            Set oTempClass = New PS_SD310
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD400") Then
            Set oTempClass = New PS_SD400
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD471") Then
            Set oTempClass = New PS_SD471
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD476") Then
            Set oTempClass = New PS_SD476
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD480") Then
            Set oTempClass = New PS_SD480
            Call oTempClass.LoadForm
        End If
        
        '//LBG
        If (pval.MenuUID = "PS_PP005") Then      '//
            Set oTempClass = New PS_PP005
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_CO010") Then '//원가요소등록
            Set oTempClass = New PS_CO010
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_CO020") Then '//원가요소그룹등록
            Set oTempClass = New PS_CO020
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_CO050") Then '//활동유형등록
            Set oTempClass = New PS_CO050
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_CO070") Then '//활동량 계산
            Set oTempClass = New PS_CO070
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_CO080") Then '//코스트센터비용집계
            Set oTempClass = New PS_CO080
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_CO090") Then '//통계주요지표값입력
            Set oTempClass = New PS_CO090
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP140") Then '//제품입고대장(휘팅-창원)
            Set oTempClass = New PS_PP140
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP180") Then '//본사휘팅출하대기현황
            Set oTempClass = New PS_PP180
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP350") Then '//부품작업일지
            Set oTempClass = New PS_PP350
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP410") Then '//지원공수현황
            Set oTempClass = New PS_PP410
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP050") Then '//작업공정내역조회
            Set oTempClass = New PS_PP050
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP170") Then '//휘팅재공품미포장현황
            Set oTempClass = New PS_PP170
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP150") Then '//재고현황
            Set oTempClass = New PS_PP150
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP320") Then '//부품주문번호별수주잔량현황
            Set oTempClass = New PS_PP320
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP150") Then '//공정별 진행(재공)현황
            Set oTempClass = New PS_PP150
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP360") Then '//공정코드 LIST
            Set oTempClass = New PS_PP360
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP300") Then '//휘팅서울포장현황
            Set oTempClass = New PS_PP300
            Call oTempClass.LoadForm
         ElseIf (pval.MenuUID = "PS_MM243") Then '//임가공원재료입고현황
            Set oTempClass = New PS_MM243
            Call oTempClass.LoadForm
         ElseIf (pval.MenuUID = "PS_PP631") Then '//PACKING LIST 미작성 검사현황
            Set oTempClass = New PS_PP631
            Call oTempClass.LoadForm
         ElseIf (pval.MenuUID = "PS_PP632") Then '//제품재고 LIST
            Set oTempClass = New PS_PP632
            Call oTempClass.LoadForm
         ElseIf (pval.MenuUID = "PS_PP634") Then '//출고현황 품목별
            Set oTempClass = New PS_PP634
            Call oTempClass.LoadForm
         ElseIf (pval.MenuUID = "PS_PP636") Then '//원재료 입고현황(멀티)
            Set oTempClass = New PS_PP636
            Call oTempClass.LoadForm
         ElseIf (pval.MenuUID = "PS_PP637") Then '//재공품 내역서(제품기준)
            Set oTempClass = New PS_PP637
            Call oTempClass.LoadForm
         ElseIf (pval.MenuUID = "PS_PP690") Then '//재공품 내역서(공정기준)
            Set oTempClass = New PS_PP690
            Call oTempClass.LoadForm
         ElseIf (pval.MenuUID = "PS_PP638") Then '//생산현황(집계표)
            Set oTempClass = New PS_PP638
            Call oTempClass.LoadForm
         ElseIf (pval.MenuUID = "PS_PP683") Then '//개인별 공정 생산량 현황(멀티)
            Set oTempClass = New PS_PP683
            Call oTempClass.LoadForm
         ElseIf (pval.MenuUID = "PS_PP684") Then '//공정별 개인 생산량 현황(멀티)
            Set oTempClass = New PS_PP684
            Call oTempClass.LoadForm
         ElseIf (pval.MenuUID = "PS_PP639") Then '//P/O별 생산현황
            Set oTempClass = New PS_PP639
            Call oTempClass.LoadForm
         ElseIf (pval.MenuUID = "PS_PP660") Then '//제품별 불량현황(멀티)
            Set oTempClass = New PS_PP660
            Call oTempClass.LoadForm
         ElseIf (pval.MenuUID = "PS_PP675") Then '//검사현황 미완료(멀티)
            Set oTempClass = New PS_PP675
            Call oTempClass.LoadForm
         ElseIf (pval.MenuUID = "PS_PP662") Then '//원인별 불량현황(멀티)
            Set oTempClass = New PS_PP662
            Call oTempClass.LoadForm
         ElseIf (pval.MenuUID = "PS_PP664") Then '//공정별 불량현황(멀티)
            Set oTempClass = New PS_PP664
            Call oTempClass.LoadForm
         ElseIf (pval.MenuUID = "PS_PP510") Then '//미완료현황(멀티)
           Set oTempClass = New PS_PP510
           Call oTempClass.LoadForm
         ElseIf (pval.MenuUID = "PS_PP520") Then '//생산실적
           Set oTempClass = New PS_PP520
           Call oTempClass.LoadForm
         ElseIf (pval.MenuUID = "PS_PP644") Then '//원재료 재고 LIST (세부내역)-멀티
           Set oTempClass = New PS_PP644
           Call oTempClass.LoadForm
         ElseIf (pval.MenuUID = "PS_PP688") Then '//생산의뢰서
           Set oTempClass = New PS_PP688
           Call oTempClass.LoadForm
         ElseIf (pval.MenuUID = "PS_PP655") Then '//수율현황
           Set oTempClass = New PS_PP655
           Call oTempClass.LoadForm
         ElseIf (pval.MenuUID = "PS_PP400") Then '//납기초과 생산품 현황
           Set oTempClass = New PS_PP400
           Call oTempClass.LoadForm
         ElseIf (pval.MenuUID = "PS_PP480") Then '//생산완료 현황
           Set oTempClass = New PS_PP480
           Call oTempClass.LoadForm
         ElseIf (pval.MenuUID = "PS_PP700") Then '//월별 생산의뢰 접수현황
           Set oTempClass = New PS_PP700
           Call oTempClass.LoadForm
         ElseIf (pval.MenuUID = "PS_MM250") Then '//구매실적가격현황
           Set oTempClass = New PS_MM250
           Call oTempClass.LoadForm
         ElseIf (pval.MenuUID = "PS_MM091") Then '//기타출고현황
           Set oTempClass = New PS_MM091
           Call oTempClass.LoadForm
         ElseIf (pval.MenuUID = "PS_QM100") Then '//외주제작 수입검사현황
           Set oTempClass = New PS_QM100
           Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP361") Then '//작지별 생산진행현황
            Set oTempClass = New PS_PP361
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP780") Then '//휘팅적정재고대장
            Set oTempClass = New PS_PP780
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP911") Then '//작번이력현황
            Set oTempClass = New PS_PP911
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_FI914") Then '//재작업현황
            Set oTempClass = New PS_FI914
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM041") Then '//품의승인조회
            Set oTempClass = New PS_MM041
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD909") Then '//Size별 휘팅 생산현황
            Set oTempClass = New PS_SD909
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM920") Then '//부자재 불출대장(일자별)
            Set oTempClass = New PS_MM920
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP310") Then '//LOT별 지시 대비 생산현황(휘팅)
            Set oTempClass = New PS_PP310
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD903") Then '//출하대비출하요청현황
            Set oTempClass = New PS_SD903
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD906") Then '//생산P/O현황
            Set oTempClass = New PS_SD906
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM923") Then '//제품재고이동현황
            Set oTempClass = New PS_MM923
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM924") Then '//제품반품현황
            Set oTempClass = New PS_MM924
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP919") Then '//부문별 공정별 생산현황
            Set oTempClass = New PS_PP919
            Call oTempClass.LoadForm
        End If

        '//NGY
        If (pval.MenuUID = "PS_PP060") Then '//작업일보외 공수등록
            Set oTempClass = New PS_PP060
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM025") Then '//품질검사 납품처 변경
            Set oTempClass = New PS_QM025
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM015") Then '//통합구매 품의 수동생성
            Set oTempClass = New PS_MM015
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SY002") Then '//프린트설정
            Set oTempClass = New PS_SY002
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM001") Then '//스크랩단가등록
            Set oTempClass = New PS_MM001
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD085") Then '//입금조회
            Set oTempClass = New PS_SD085
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD050") Then '//기성매출등록
            Set oTempClass = New PS_SD050
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM070") Then '//불합리적출등록
            Set oTempClass = New PS_QM070
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP081") Then '//생산완료등록(방산부품)
            Set oTempClass = New PS_PP081
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM060") Then '//품질검사일지등록
            Set oTempClass = New PS_QM060
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_CO501") Then '//품목별 원가입력
            Set oTempClass = New PS_CO501
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_CO502") Then '//평균 원가항목입력
            Set oTempClass = New PS_CO502
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_CO503") Then '//일일 가득액 및 생산실적 집계
            Set oTempClass = New PS_CO503
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP110") Then '//추가공수(비용)수정
            Set oTempClass = New PS_PP110
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_CO504") Then '//일일판매생산현황
            Set oTempClass = New PS_CO504
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM173") Then '//지체상금현황
            Set oTempClass = New PS_MM173
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM260") Then '//외주재공(절단)품 미출고대장
            Set oTempClass = New PS_MM260
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM220") Then '//외주 스크랩 수불부
            Set oTempClass = New PS_MM220
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP370") Then '//휘팅 제품생산 및 판매현황
            Set oTempClass = New PS_PP370
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM270") Then '//휘팅 외주 미입고현황
            Set oTempClass = New PS_MM270
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP790") Then '//휘팅 결품현황
            Set oTempClass = New PS_PP790
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD150") Then '//출하확인(요청대비)
            Set oTempClass = New PS_SD150
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM218") Then '//장기재고대장
            Set oTempClass = New PS_MM218
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP640") Then '//거래처LOTNO별현황(멀티)
            Set oTempClass = New PS_PP640
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP380") Then '//작업지시서미발행내역대장
            Set oTempClass = New PS_PP380
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP390") Then '//원재료불출현황
            Set oTempClass = New PS_PP390
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP007") Then '//공정장비등록
            Set oTempClass = New PS_PP007
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP914") Then '//공정장비별가동율현황
            Set oTempClass = New PS_PP914
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP810") Then '//TOOL사용실적
            Set oTempClass = New PS_PP810
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD490") Then '//휘팅출하요청대비출하현황(월별)
            Set oTempClass = New PS_SD490
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD901") Then '//기계 판매오더 수주처 변경처리
            Set oTempClass = New PS_SD901
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP710") Then '//휘팅이동등록대장
            Set oTempClass = New PS_PP710
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP930") Then '//일일생산계획등록
            Set oTempClass = New PS_PP930
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD410") Then '//제품재고이동등록대장
            Set oTempClass = New PS_SD410
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD420") Then '//(구로)미출하대장
            Set oTempClass = New PS_SD420
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP720") Then '//휘팅단중비교현황
            Set oTempClass = New PS_PP720
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP330") Then '//PO별생산현황
            Set oTempClass = New PS_PP330
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM002") Then '//제품BOM등록관리
            Set oTempClass = New PS_MM002
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM280") Then '//기간별SCRAP반납내역(멀티)
            Set oTempClass = New PS_MM280
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP820") Then '//판매목표및생산진도관리(부품)
            Set oTempClass = New PS_PP820
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM290") Then '//제품BOM대장(포장)
            Set oTempClass = New PS_MM290
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_CO210") Then '//휘팅제품원가계산
            Set oTempClass = New PS_CO210
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP830") Then '//M/G생산스크랩수불부
            Set oTempClass = New PS_PP830
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP840") Then '//M/G원재료반품현황
            Set oTempClass = New PS_PP840
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP850") Then '//M/G원재료반품현황
            Set oTempClass = New PS_PP850
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SM030") Then '//M/G원재료반품현황
            Set oTempClass = New PS_SM030
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP130") Then '//생산설비등록
            Set oTempClass = New PS_PP130
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP860") Then '//월생산계획대비실적조회
            Set oTempClass = New PS_PP860
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP160") Then '//설비가동시간등록
            Set oTempClass = New PS_PP160
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP870") Then '//가동분석현황
            Set oTempClass = New PS_PP870
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM903") Then '//외부실패비용등록
            Set oTempClass = New PS_QM903
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_FI210") Then '//원가요소(배부담당)별비용집계대장
            Set oTempClass = New PS_FI210
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP420") Then '//개인별비가동내역집계대장
            Set oTempClass = New PS_PP420
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM101") Then '//기계공구류검사일지등록
            Set oTempClass = New PS_QM101
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM102") Then '//기계공구류수입검사일지등록
            Set oTempClass = New PS_QM102
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_HR200") Then '//인사급여코드등록
            Set oTempClass = New PS_HR200
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_HR300") Then '//경영방침등록
            Set oTempClass = New PS_HR300
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_HR305") Then '//사업부KPI등록
            Set oTempClass = New PS_HR305
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_HR310") Then '//팀KPI등록
            Set oTempClass = New PS_HR310
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_HR315") Then '//담당KPI등록
            Set oTempClass = New PS_HR315
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_HR320") Then '//담당KPI등록
            Set oTempClass = New PS_HR320
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_HR325") Then '//KPI목표등록마감
            Set oTempClass = New PS_HR325
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM120") Then '//MG생산불량실패비용
            Set oTempClass = New PS_QM120
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM130") Then '//픔질검사현황
            Set oTempClass = New PS_QM130
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM140") Then '//픔질실패비용현황
            Set oTempClass = New PS_QM140
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD510") Then '//기성매출현황
            Set oTempClass = New PS_SD510
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP036") Then '//MG생산 일보조회(수정)
            Set oTempClass = New PS_PP036
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP880") Then '//담당별생산계획대비실적대장
            Set oTempClass = New PS_PP880
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD520") Then '//기성매출제품입고현황
            Set oTempClass = New PS_SD520
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP111") Then '//추가예상공수(비용)등록대상
            Set oTempClass = New PS_PP111
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD241") Then '//미납품현황(포장용)
            Set oTempClass = New PS_SD241
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM945") Then '//AP송장미발행 현황
            Set oTempClass = New PS_MM945
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP890") Then '//외주가공목표대비실적현황
            Set oTempClass = New PS_PP890
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_HR385") Then '//사업부KPI목표출력
            Set oTempClass = New PS_HR385
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP910") Then '//담당별생산/판매현황
            Set oTempClass = New PS_PP910
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_HR370") Then '//KPI담당실적현황
            Set oTempClass = New PS_HR370
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_HR380") Then '//KPI팀실적현황
            Set oTempClass = New PS_HR380
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_HR395") Then '//KPI월별실적현황
            Set oTempClass = New PS_HR395
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_HR365") Then '//담당KPI목표출력
            Set oTempClass = New PS_HR365
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_HR375") Then '//팀KPI목표출력
            Set oTempClass = New PS_HR375
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM220") Then '//M/G검사성적조회
            Set oTempClass = New PS_QM220
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM045") Then '//통계적품질분석결과
            Set oTempClass = New PS_QM045
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD190") Then '//판매실적조회
            Set oTempClass = New PS_SD190
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM240") Then '//계측기등록
            Set oTempClass = New PS_QM240
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM241") Then '//계측기이력등록
            Set oTempClass = New PS_QM241
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM242") Then '//계측기이력조회
            Set oTempClass = New PS_QM242
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP190") Then '//멀티금형관리
            Set oTempClass = New PS_PP190
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP051") Then '//M/G생산작업지시발행현황
            Set oTempClass = New PS_PP051
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM946") Then '//부자재입고준수확인(KPI)
            Set oTempClass = New PS_MM946
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM041") Then '//검사성적서(신양식)
            Set oTempClass = New PS_QM041
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM190") Then '//외주가공비현황
            Set oTempClass = New PS_MM190
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM080") Then '//일일품질현황
            Set oTempClass = New PS_QM080
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM090") Then '//제품별불량현황
            Set oTempClass = New PS_QM090
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM300") Then '//구매청구현황
            Set oTempClass = New PS_MM300
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP191") Then '//멀티금형조회
            Set oTempClass = New PS_PP191
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM091") Then '//연간제품별불량현황
            Set oTempClass = New PS_QM091
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_HR400") Then '//전문직 평가마스터 입력
            Set oTempClass = New PS_HR400
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_HR401") Then '//전문직 평가착안점 마스터 입력
            Set oTempClass = New PS_HR401
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_HR402") Then '//전문직 평가 배분표등록
            Set oTempClass = New PS_HR402
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_HR403") Then '//서약서
            Set oTempClass = New PS_HR403
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_HR405") Then '//전문직 인사평가 참고사항입력
            Set oTempClass = New PS_HR405
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_HR406") Then '//전문직 인사평가 참고사항입력
            Set oTempClass = New PS_HR406
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_HR407") Then '//전문직 인사평가 이력조회
            Set oTempClass = New PS_HR407
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_HR409") Then '//비밀번호변경
            Set oTempClass = New PS_HR409
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_HR410") Then '//전문직 평가대상자등록
            Set oTempClass = New PS_HR410
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_HR412") Then '//전문직 평가대상자 그룹핑
            Set oTempClass = New PS_HR412
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_HR416") Then '//전문직 평가
            Set oTempClass = New PS_HR416
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_HR418") Then '//전문직 평가
            Set oTempClass = New PS_HR418
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM012") Then '//자재 순환품관리
            Set oTempClass = New PS_MM012
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_FX005") Then '//고정자산등록
            Set oTempClass = New PS_FX005
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_FX010") Then '//고정자산이력등록
            Set oTempClass = New PS_FX010
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_FX020") Then '//감가상각등록
            Set oTempClass = New PS_FX020
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_FX025") Then '//감가상각 분개처리
            Set oTempClass = New PS_FX025
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_FX210") Then '//고정자산이력현황
            Set oTempClass = New PS_FX210
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_FX220") Then '//자본적지출현황
            Set oTempClass = New PS_FX220
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_FX230") Then '//고정자산구입현황
            Set oTempClass = New PS_FX230
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_FX240") Then '//고정자산현황
            Set oTempClass = New PS_FX240
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM081") Then '//방산부품 검사사양서등록
            Set oTempClass = New PS_QM081
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM082") Then '//방산부품 원재료입고 등록
            Set oTempClass = New PS_QM082
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM083") Then '//방산부품 출고검사 등록
            Set oTempClass = New PS_QM083
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM247") Then '//매입집계 및 세금계산서 접수 현황
            Set oTempClass = New PS_MM247
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM245") Then '//방산절단품입출고현황
            Set oTempClass = New PS_QM245
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM250") Then '//방산절단품시험분석결과
            Set oTempClass = New PS_QM250
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM255") Then '//방산소재입고현황
            Set oTempClass = New PS_QM255
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP960") Then '//부품일자별생산현황
            Set oTempClass = New PS_PP960
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_HR420") Then '//전문직평가 그룹리스트 조회
            Set oTempClass = New PS_HR420
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP965") Then '//부품일자별판매현황
            Set oTempClass = New PS_PP965
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_CO160") Then '//재공원가 변경(제품변경)
            Set oTempClass = New PS_CO160
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP015") Then '//포장생산 작업지시서 발행
            Set oTempClass = New PS_PP015
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP970") Then '//생산팀O/T현황
            Set oTempClass = New PS_PP970
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP043") Then '//포장사업팀 생산일보등록(공정)
            Set oTempClass = New PS_PP043
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM006") Then '//구매요청 승인
            Set oTempClass = New PS_MM006
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM135") Then '//외주반출등록(포장사업팀)
            Set oTempClass = New PS_MM135
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM138") Then '//외주입고등록(포장사업팀)
            Set oTempClass = New PS_MM138
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM096") Then '//원재료불출등록(포장사업팀)
            Set oTempClass = New PS_MM096
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM955") Then '//외주수불부(포장)
            Set oTempClass = New PS_MM955
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP600") Then '//재공현황(포장)
            Set oTempClass = New PS_PP600
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM260") Then '//검사게이지관리등록
            Set oTempClass = New PS_QM260
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP975") Then '// M/G SHIPPING MARK출력
            Set oTempClass = New PS_PP975
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM261") Then '//검사게이지이력조회
            Set oTempClass = New PS_QM261
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP046") Then '//작업일보조회(포장사업팀)
            Set oTempClass = New PS_PP046
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM960") Then '//외주입고현황(포장)
            Set oTempClass = New PS_MM960
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD485") Then '//(포장)매출실적현황
            Set oTempClass = New PS_SD485
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_HR404") Then '//전문직평가 가감항목마스터등록
            Set oTempClass = New PS_HR404
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_HR413") Then '//전문직평가 정량평가등록
            Set oTempClass = New PS_HR413
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_HR414") Then '//전문직평가 가감항목등록
            Set oTempClass = New PS_HR414
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_HR421") Then '//전문직평가 조회(관리자용)
            Set oTempClass = New PS_HR421
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_HR415") Then '//전문직평가 정량평가 결과등록
            Set oTempClass = New PS_HR415
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_HR422") Then '//전문직인사평가세부내역
            Set oTempClass = New PS_HR422
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_HR423") Then '//전문직정량평가현황
            Set oTempClass = New PS_HR423
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_HR419") Then '//전문직평가 등급처리
            Set oTempClass = New PS_HR419
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM097") Then
            Set oTempClass = New PS_MM097
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_HR424") Then '//전문직 평가결과 조회
            Set oTempClass = New PS_HR424
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP355") Then '//부품생산성지표
            Set oTempClass = New PS_PP355
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_FI960") Then '//입,퇴사자현황
            Set oTempClass = New PS_FI960
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_FTP") Then '//FTP전송
            Set oTempClass = New PS_FTP
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP083") Then '//포장생산팀 생산완료등록
            Set oTempClass = New PS_PP083
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM034") Then '//M/G원소재수입검사기준등록
            Set oTempClass = New PS_QM034
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM035") Then '//M/G원소재수입검자자료등록
            Set oTempClass = New PS_QM035
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD610") Then '//내부거래출고현황
            Set oTempClass = New PS_SD610
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP095") Then '//M/G거래명세표등록
            Set oTempClass = New PS_PP095
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM240") Then '//부품일자별외주수불현황
            Set oTempClass = New PS_MM240
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP037") Then '//포장생산 일보조회(수정)
            Set oTempClass = New PS_PP037
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM255") Then '//구매실적가격조회
            Set oTempClass = New PS_MM255
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP961") Then '//쇼링생산현황
            Set oTempClass = New PS_PP961
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD486") Then '//품목별판매현황
            Set oTempClass = New PS_SD486
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM310") Then '//M/G원소재수입검사현황
            Set oTempClass = New PS_QM310
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM153") Then '//외주입고 품질검사
            Set oTempClass = New PS_MM153
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP200") Then '//포장공정packing List
            Set oTempClass = New PS_PP200
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP210") Then '//패킹번호별재고대장
            Set oTempClass = New PS_PP210
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM021") Then '//MG품질검사 1차승인
            Set oTempClass = New PS_QM021
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM022") Then '//MG품질검사 2차승인
            Set oTempClass = New PS_QM022
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP215") Then '//원재료 월 소요량 등록
            Set oTempClass = New PS_PP215
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_FX250") Then '//고정자산이력현황(개별)
            Set oTempClass = New PS_FX250
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP220") Then '//월생산계획및원자재공급현황
            Set oTempClass = New PS_PP220
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP351") Then '//개인작업시간 현황
            Set oTempClass = New PS_PP351
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP601") Then '//포장 소요량대비원재료불출현황
            Set oTempClass = New PS_PP601
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_CO250") Then '//개인별 퇴충계산
            Set oTempClass = New PS_CO250
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP096") Then '//M/G출고계획
            Set oTempClass = New PS_PP096
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM133") Then '//외주반출승인
            Set oTempClass = New PS_MM133
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM008") Then '//제품공정단중등록
            Set oTempClass = New PS_MM008
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM916") Then '//외주가공스크랩발생현황
            Set oTempClass = New PS_MM916
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM055") Then '//부품 제품검사사양 등록
            Set oTempClass = New PS_QM055
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_CO260") Then '//연차설정등록
            Set oTempClass = New PS_CO260
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM056") Then '//부품-재품검사서 등록
            Set oTempClass = New PS_QM056
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM270") Then '//방산부품불량현황
            Set oTempClass = New PS_QM270
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_CO131") Then '//원가계산재공현황
            Set oTempClass = New PS_CO131
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM150") Then '//제안시상기준표등록
            Set oTempClass = New PS_QM150
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM151") Then '//제안등록
            Set oTempClass = New PS_QM151
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM152") Then '//제안평가(QC)
            Set oTempClass = New PS_QM152
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM160") Then '//제안현황
            Set oTempClass = New PS_QM160
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM161") Then '//제안개인별집계분석
            Set oTempClass = New PS_QM161
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM003") Then '//화물운송가격표등록
            Set oTempClass = New PS_MM003
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM004") Then '//운송확인증 출력
            Set oTempClass = New PS_MM004
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM195") Then '//운송비지출내역서
            Set oTempClass = New PS_MM195
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_FX260") Then '//자산LABEL출력
            Set oTempClass = New PS_FX260
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP230") Then '//작업실적입력현황(현장)
            Set oTempClass = New PS_PP230
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM930") Then '//거래처별년간매입집계대장
            Set oTempClass = New PS_MM930
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_CO510") Then '//원가계산사전점검조회
            Set oTempClass = New PS_CO510
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_CO520") Then '//제품생산 원가항목별 조회
            Set oTempClass = New PS_CO520
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_GA901") Then '//방문자 정보등록
            Set oTempClass = New PS_GA901
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_GA910") Then '//출입예정자등록
            Set oTempClass = New PS_GA910
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_GA915") Then '//출입예정자접수
            Set oTempClass = New PS_GA915
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_GA920") Then '//출입자 입문처리
            Set oTempClass = New PS_GA920
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_GA925") Then '//출문 등록
            Set oTempClass = New PS_GA925
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_GA930") Then '//차량계근등록
            Set oTempClass = New PS_GA930
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_GA960") Then '//출입예정자 조회
            Set oTempClass = New PS_GA960
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_GA962") Then '//일일방문객 조회
            Set oTempClass = New PS_GA962
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_GA964") Then '//차량계근기록 조회
            Set oTempClass = New PS_GA964
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_GA966") Then '//제품반출일지
            Set oTempClass = New PS_GA966
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD005") Then '//거래처별제품단가등록(분말)
            Set oTempClass = New PS_SD005
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD006") Then '//판매계획등록(분말)
            Set oTempClass = New PS_SD006
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_GA968") Then '//원자재반입일지
            Set oTempClass = New PS_GA968
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_GA970") Then '//부자재반입일지
            Set oTempClass = New PS_GA970
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM110") Then '//분말 원재료 검사사양서등록
            Set oTempClass = New PS_QM110
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM115") Then '//분말 원재료 수입검사등록
            Set oTempClass = New PS_QM115
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP025") Then '//분말 생산계획등록
            Set oTempClass = New PS_PP025
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP084") Then '//분말 생산실적등록
            Set oTempClass = New PS_PP084
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP092") Then '//분말팩킹관리
            Set oTempClass = New PS_PP092
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM007") Then '//분말제품검사 사양서
            Set oTempClass = New PS_QM007
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM008") Then '//분말제품검사 사양서
            Set oTempClass = New PS_QM008
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PACKING_PD") Then '//분말포장지시등록
            Set oTempClass = New PS_PACKING_PD
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP048") Then '//분말스크랩등록
            Set oTempClass = New PS_PP048
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM100") Then '//부자재 불출등록
            Set oTempClass = New PS_MM100
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP325") Then '//부품주문번호별출고대장(일자별)
            Set oTempClass = New PS_PP325
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_FI190") Then '//보조원장조회(EXCEL다운)
            Set oTempClass = New PS_FI190
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SM040") Then '//재고전기리스트
            Set oTempClass = New PS_SM040
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP155") Then '//분말생산공정진행현황
            Set oTempClass = New PS_PP155
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD285") Then '//타계정출고현황(샘플,감모손)
            Set oTempClass = New PS_SD285
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_CO170") Then '//재공 완료대체품 관리등록(분말 제품 통합관리)
            Set oTempClass = New PS_CO170
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP730") Then '//분말일자별작업조회
            Set oTempClass = New PS_PP730
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP021") Then '//작업일보 작업자 분류등록
            Set oTempClass = New PS_PP021
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP044") Then '//부품작업일보등록
            Set oTempClass = New PS_PP044
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP686") Then '//멀티생산입고대기(재공)현황
            Set oTempClass = New PS_PP686
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM007") Then '//통구실 품목코드생성 요청등록
            Set oTempClass = New PS_MM007
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM921") Then '//방산부품용 부자재불출현황
            Set oTempClass = New PS_MM921
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_FX241") Then '//방산부품용 고정자산현황
            Set oTempClass = New PS_FX241
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD960") Then '//거래처별월별채권현황
            Set oTempClass = New PS_SD960
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM510") Then '//코드생성요청진행현황
            Set oTempClass = New PS_MM510
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP047") Then '//금속분말 재작업등록
            Set oTempClass = New PS_PP047
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM919") Then '//부서별품목별실적집계조회
            Set oTempClass = New PS_MM919
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM460") Then '//금속분말품질지표
            Set oTempClass = New PS_QM460
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP995") Then '//분말LABEL출력
            Set oTempClass = New PS_PP995
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP996") Then '//분말LABEL출력2
            Set oTempClass = New PS_PP996
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM200") Then '//수불명세서등록(년)
            Set oTempClass = New PS_MM200
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM201") Then '//수불명세서(연간)
            Set oTempClass = New PS_MM201
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM205") Then '//재공수불등록(년)
            Set oTempClass = New PS_MM205
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM206") Then '//재공수불출력
            Set oTempClass = New PS_MM206
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM208") Then '//스크랩회수운송지출내역서
            Set oTempClass = New PS_MM208
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM011") Then '//고객불만등록
            Set oTempClass = New PS_QM011
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM012") Then '//고객불만관리대장
            Set oTempClass = New PS_QM012
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP642") Then '//M/G생산기간별생산실적현황
            Set oTempClass = New PS_PP642
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP643") Then '//기간별제품기타출고현황
            Set oTempClass = New PS_PP643
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM470") Then '//분말제품검사사양서출력
            Set oTempClass = New PS_QM470
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP835") Then '//M/G스크랩발생현황
            Set oTempClass = New PS_PP835
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM935") Then '//금속분말샘플타계정출고진행현황
            Set oTempClass = New PS_MM935
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP011") Then '//거래처제품(연결)코드등록
            Set oTempClass = New PS_PP011
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP845") Then '//M/G생산원재료기타출고현황
            Set oTempClass = New PS_PP845
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD477") Then '//품목별판매집계현황
            Set oTempClass = New PS_SD477
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM017") Then '//통합구매미등록거래처조회
            Set oTempClass = New PS_MM017
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP151") Then '//공정별진행현황(부품,기초)
            Set oTempClass = New PS_PP151
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM020") Then '//자재순환품조회
            Set oTempClass = New PS_MM020
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_GA105") Then '//일자별 무재해 일수조회
            Set oTempClass = New PS_GA105
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM025") Then '//반출LABEL출력
            Set oTempClass = New PS_MM025
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM092") Then '//분말생산 샘플출고등록
            Set oTempClass = New PS_MM092
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM065") Then '//포장사업팀 제품검사항목등록
            Set oTempClass = New PS_QM065
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM066") Then '//포장사업팀 제품검사서등록
            Set oTempClass = New PS_QM066
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP152") Then '//공정별 생산현황
            Set oTempClass = New PS_PP152
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM067") Then '//KPI 납기준수율 안강 황영수
            Set oTempClass = New PS_QM067
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "ZPS_MGSCRAP") Then '//MG스크랩수불부
            Set oTempClass = New ZPS_MGSCRAP
            Call oTempClass.LoadForm
        '//HYS
        ElseIf (pval.MenuUID = "PS_PP645") Then '//출고현황 품목별 포장사업부
            Set oTempClass = New PS_PP645
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_CO185") Then '//담당별 판매/원재료/R&M 단가
            Set oTempClass = New PS_CO185
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM225") Then '//담당별 판매/원재료/R&M 단가
            Set oTempClass = New PS_MM225
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP012") Then '//분말 표준량 / 분류코드 등록
            Set oTempClass = New PS_PP012
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD011") Then '//거래처별 납품 라벨 등록
            Set oTempClass = New PS_SD011
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP097") Then '//검사사양서 확인 검증
            Set oTempClass = New PS_PP097
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_CO530") Then '//원가_분말부자재비용분석
            Set oTempClass = New PS_CO530
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_CO540") Then '//원가_분말부자재비용분석
            Set oTempClass = New PS_CO540
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_FX270") Then '//고정자산관리대장(개별)
            Set oTempClass = New PS_FX270
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM320") Then '//M/G 계량치 분석값 등록
            Set oTempClass = New PS_QM320
            Call oTempClass.LoadForm
        End If
        
        '//SMG
        If (pval.MenuUID = "PS_PP055") Then '//스크랩 원재료 불량 반납
            Set oTempClass = New PS_PP055
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD295") Then '거래처별 납기 분석표
            Set oTempClass = New PS_SD295
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP945") Then '일일근무계획등록
            Set oTempClass = New PS_PP945
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP940") Then '일일생산계획등록
            Set oTempClass = New PS_PP940
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD330") Then '품목별 월별 판매현황 (2011.10.17 추가)
            Set oTempClass = New PS_SD330
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM940") Then '품의발주 대비 가입고 현황 (2011.10.26 추가)
            Set oTempClass = New PS_MM940
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_HR000") Then '사규위반 조치 기준[항목]등록 (2011.11.10 추가)
            Set oTempClass = New PS_HR000
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_HR001") Then '사규위반 조치 기준[세부사항]등록 (2011.11.11 추가)
            Set oTempClass = New PS_HR001
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_HR002") Then '사규위반 조치 기준[세세부사항]등록 (2011.11.11 추가)
            Set oTempClass = New PS_HR002
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_HR100") Then '사규위반 조치등록 (2011.11.14 추가)
            Set oTempClass = New PS_HR100
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_HR105") Then '사규위반조치등록조회 (2012.08.06 추가)
            Set oTempClass = New PS_HR105
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_AD100") Then '프로그램 수정 및 추가 요청 사항 등록 (2011.12.14 추가)
            Set oTempClass = New PS_AD100
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_AD101") Then '프로그램 수정 및 추가 요청 사항 접수 (2011.12.19 추가)
            Set oTempClass = New PS_AD101
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_AD102") Then '프로그램 수정 및 추가 요청 사항 완료 (2011.12.19 추가)
            Set oTempClass = New PS_AD102
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_AD900") Then '프로그램 수정 및 추가 요청 현황 (2011.12.30 추가)
            Set oTempClass = New PS_AD900
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP900") Then '품목별 공정 진행현황 (2011.12.20 추가)
            Set oTempClass = New PS_PP900
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP362") Then '작번별 생산진행현황2(2012.01.28 추가)
            Set oTempClass = New PS_PP362
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_HR330") Then '담당원KPI실적등록(2012.02.28 추가)
            Set oTempClass = New PS_HR330
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_HR335") Then '담당KPI실적등록(2012.03.06 추가)
            Set oTempClass = New PS_HR335
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_HR340") Then '팀KPI실적등록(2012.03.07 추가)
            Set oTempClass = New PS_HR340
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_HR345") Then '사업부KPI실적등록(2012.03.08 추가)
            Set oTempClass = New PS_HR345
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_HR350") Then 'KPI실적마감(2012.03.30 추가)
            Set oTempClass = New PS_HR350
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM285") Then '수명제한 품목 관리대장(2012.03.27 추가)
            Set oTempClass = New PS_MM285
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP031") Then '품목별 공수 조회(2012.04.03 추가)
            Set oTempClass = New PS_PP031
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP032") Then '표준공수등록(2012.04.10 추가)
            Set oTempClass = New PS_PP032
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP905") Then '생산부하현황(2012.04.18 추가)
            Set oTempClass = New PS_PP905
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM200") Then '표준등록(2012.05.07 추가)
            Set oTempClass = New PS_QM200
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM210") Then '사내표준등록현황(2012.05.08 추가)
            Set oTempClass = New PS_QM210
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM215") Then '사내표준관리대장(2012.05.08 추가)
            Set oTempClass = New PS_QM215
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM230") Then '엔드베어링 품질경향 등록(2012.05.22 추가)
            Set oTempClass = New PS_QM230
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP906") Then '개인별작업시간현황(2012.06.12 추가)
            Set oTempClass = New PS_PP906
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM947") Then '공정별 외주가공비 현황(2012.06.22 추가)
            Set oTempClass = New PS_MM947
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP908") Then '생산업무현황(2012.06.29 추가)
            Set oTempClass = New PS_PP908
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_GA005") Then '정전이력등록(2012.07.13 추가)
            Set oTempClass = New PS_GA005
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_GA010") Then '정전이력조회(2012.07.18 추가)
            Set oTempClass = New PS_GA010
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP909") Then '설계 실동공수 등록현황(기계)(2012.08.06 추가)
            Set oTempClass = New PS_PP909
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_FI950") Then '손익계산서(2012.08.09 추가)
            Set oTempClass = New PS_FI950
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_FI600") Then '손익계산서 과목관리(2012.08.10 추가)
            Set oTempClass = New PS_FI600
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD950") Then '채권 만기일 관리(2012.08.21 추가)
            Set oTempClass = New PS_SD950
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_GA050") Then '총무관리코드등록(2012.08.30 추가)
            Set oTempClass = New PS_GA050
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_GA052") Then '검진대상자등록(2012.08.28 추가)
            Set oTempClass = New PS_GA052
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_GA053") Then '검진대상자조회(2012.08.31 추가)
            Set oTempClass = New PS_GA053
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_GA054") Then '건강관리이력등록(2012.09.03 추가)
            Set oTempClass = New PS_GA054
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_GA055") Then '건강관리이력조회(2012.09.03 추가)
            Set oTempClass = New PS_GA055
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_GA056") Then '유소견자등록(2012.09.04 추가)
            Set oTempClass = New PS_GA056
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_GA057") Then '유소견자조회(2012.09.06 추가)
            Set oTempClass = New PS_GA057
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_GA100") Then '무재해목표등록(2012.09.12 추가)
            Set oTempClass = New PS_GA100
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_GA101") Then '무재해목표조회(2012.09.14 추가)
            Set oTempClass = New PS_GA101
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_GA103") Then '무재해진행실적조회(2012.09.14 추가)
            Set oTempClass = New PS_GA103
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_GA150") Then '안전보호구지급등록(2012.09.18 추가)
            Set oTempClass = New PS_GA150
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_GA151") Then '안전보호구지급조회(2012.09.18 추가)
            Set oTempClass = New PS_GA151
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_FI215") Then '지급현황(2012.09.28 추가)
            Set oTempClass = New PS_FI215
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM950") Then '담당별 품의구분별 구매 내역(2012.10.08 추가)
            Set oTempClass = New PS_MM950
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_GA160") Then '전산장비등록(기본정보)(2012.10.25 추가)
            Set oTempClass = New PS_GA160
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM300") Then '멀티게이지 품질부적합 등록(2012.12.21 추가)
            Set oTempClass = New PS_QM300
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM301") Then '멀티게이지 품질부적합 등록 해제(1차)(2013.01.03 추가)
            Set oTempClass = New PS_QM301
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM302") Then '멀티게이지 품질부적합 등록 해제(2차)(2013.01.07 추가)
            Set oTempClass = New PS_QM302
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM303") Then '멀티게이지 품질부적합 등록 현황(2013.01.10 추가)
            Set oTempClass = New PS_QM303
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD600") Then '견적서등록(2013.02.08 추가)
            Set oTempClass = New PS_SD600
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM951") Then '통합구매 현황 조회(2013.02.18 추가)
            Set oTempClass = New PS_MM951
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_FI421") Then '입금표 등록(2013.04.23 추가)
            Set oTempClass = New PS_FI421
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP009") Then '공정인원등록(2013.05.29 추가)
            Set oTempClass = New PS_PP009
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP980") Then '납기준수율조회(2013.06.17 추가)
            Set oTempClass = New PS_PP980
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP981") Then '가동율조회(2013.06.22 추가)
            Set oTempClass = New PS_PP981
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP983") Then '생산현황변동조회(2013.07.10 추가)
            Set oTempClass = New PS_PP983
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM141") Then '품질실패비용현황(2013.08.29 추가)
            Set oTempClass = New PS_QM141
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_FI961") Then '월별 계정별 비용현황(2013.10.03 추가)
            Set oTempClass = New PS_FI961
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP984") Then '자체가공비 일괄 조회(2013.11.11 추가)
            Set oTempClass = New PS_PP984
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_GA020") Then '출입예정자등록(2013.11.18 추가)
            Set oTempClass = New PS_GA020
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_GA021") Then '출입예정자조회및출력(2013.12.02 추가)
            Set oTempClass = New PS_GA021
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_GA022") Then '일일방문객조회및출력(2013.12.03 추가)
            Set oTempClass = New PS_GA022
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_GA023") Then '차량계근기록조회및출력(2013.12.04 추가)
            Set oTempClass = New PS_GA023
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD020") Then '작번별담당자등록(2013.12.11 추가)
            Set oTempClass = New PS_SD020
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD021") Then '판매계획등록(2013.12.11 추가)
            Set oTempClass = New PS_SD021
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD022") Then '판매계획대비실적현황(2013.12.13 추가)
            Set oTempClass = New PS_SD022
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM400") Then '검사기준등록(2013.12.26 추가)
            Set oTempClass = New PS_QM400
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM410") Then '품질검사등록(2014.01.13 추가)
            Set oTempClass = New PS_QM410
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM430") Then '품질검사등록(2014.01.16 추가)
            Set oTempClass = New PS_QM430
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP985") Then '작번별 진행금액 일괄조회(2014.01.02 추가)
            Set oTempClass = New PS_PP985
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM965") Then '구매현황조회(2014.02.20 추가)
            Set oTempClass = New PS_MM965
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM450") Then '검사여부등록(2014.03.17 추가)
            Set oTempClass = New PS_QM450
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP988") Then '작번별 비용명세[결산/계산/완료] (2014.03.19 추가)
            Set oTempClass = New PS_PP988
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP421") Then '인원별 부서별 비가동내역조회 (2014.03.21 추가)
            Set oTempClass = New PS_PP421
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM131") Then '출고일대비검사일조회 (2014.03.21 추가)
            Set oTempClass = New PS_QM131
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_GA030") Then '예산 등록 (2014.03.28 추가)
            Set oTempClass = New PS_GA030
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_GA031") Then '실적조회 및 출력 (2014.03.31 추가)
            Set oTempClass = New PS_GA031
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM500") Then '제품검사확인(자재담당)(2014.04.01 추가)
            Set oTempClass = New PS_MM500
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD035") Then '출하리스트 등록 및 조회(2014.04.22 추가)
            Set oTempClass = New PS_SD035
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM286") Then '수기출고 대비 정상출고 현황(2014.05.13 추가)
            Set oTempClass = New PS_MM286
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD051") Then '수주상세금액등록(2014.05.21 추가)
            Set oTempClass = New PS_SD051
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD052") Then '작번별구매현황조회(2014.05.23 추가)
            Set oTempClass = New PS_SD052
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD053") Then '작번별미품의금액등록(2014.05.27 추가)
            Set oTempClass = New PS_SD053
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD054") Then '작번별추가예상공수등록(2014.05.28 추가)
            Set oTempClass = New PS_SD054
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD055") Then '작번별추가예상공수등록(2014.05.29 추가)
            Set oTempClass = New PS_SD055
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD056") Then '기성매출현황조회(2014.06.02 추가)
            Set oTempClass = New PS_SD056
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD057") Then '작번별미청구금액등록(2014.06.10 추가)
            Set oTempClass = New PS_SD057
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD058") Then '외주제작진도관리대상업체조회(2014.06.23 추가)
            Set oTempClass = New PS_SD058
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP038") Then '작업지시-투입자재추가등록,수정,삭제(2014.07.18 추가)
            Set oTempClass = New PS_PP038
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP989") Then '기간별 표준공수 대비 실동공수 조회(2014.08.25 추가)
            Set oTempClass = New PS_PP989
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_CO901") Then '월별 이익률 분석(2014.09.05 추가)
            Set oTempClass = New PS_CO901
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP915") Then '설비종합효율-기초정보조회(2014.09.26 추가)
            Set oTempClass = New PS_PP915
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_CO920") Then '월별 판매실적 대비 공수 조회(2014.09.29 송명규 추가)
            Set oTempClass = New PS_CO920
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_GA167") Then '전산장비조회(2014.10.06 송명규 추가)
            Set oTempClass = New PS_GA167
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_CO921") Then '2014.10.02_월별 작번별 수주금액 대비 공수 현황(2014.10.15 송명규 추가)
            Set oTempClass = New PS_CO921
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SY003") Then '사용자별 권한 조회
            Set oTempClass = New PS_SY003
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP916") Then '설비종합효율-집계조회(2014.10.30 추가)
            Set oTempClass = New PS_PP916
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM055") Then '미입고품 입고예정일 관리(2014.11.14 추가)
            Set oTempClass = New PS_MM055
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP039") Then '작업지시-공정추가등록,수정,삭제(2014.12.09 추가)
            Set oTempClass = New PS_PP039
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP062") Then '설비별정비현황(2015.01.29 추가)
            Set oTempClass = New PS_PP062
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP063") Then '설비별부자재사용현황(2015.02.03 추가)
            Set oTempClass = New PS_PP063
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP064") Then '(2015.02.03 추가)
            Set oTempClass = New PS_PP064
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP065") Then '(2015.02.06 추가)
            Set oTempClass = New PS_PP065
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD059") Then '(2015.02.28 추가)
            Set oTempClass = New PS_SD059
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD060") Then '(2015.02.28 추가)
            Set oTempClass = New PS_SD060
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP550") Then '바코드 리스트 출력(2015.03.10 추가)
            Set oTempClass = New PS_PP550
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP551") Then '공정상태 조회(2015.03.19 추가)
            Set oTempClass = New PS_PP551
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP552") Then '공정상태 관리(2015.03.30 추가)
            Set oTempClass = New PS_PP552
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_GA058") Then '종합건강진단 대상자 조회(2015.04.13 추가)
            Set oTempClass = New PS_GA058
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP553") Then '공정진행현황 조회(2015.04.17 추가)
            Set oTempClass = New PS_PP553
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD602") Then '견적등록(2015.07.22 추가)
            Set oTempClass = New PS_SD602
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD603") Then '견적조회(2015.12.28 추가)
            Set oTempClass = New PS_SD603
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_CO922") Then '.장비 작번별 계획금액 대비 원가현황(리스트)(2015.09.15 추가)
            Set oTempClass = New PS_CO922
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_CO923") Then '.장비 작번별 계획금액 대비 원가현황(세부정보)(2015.09.30 추가)
            Set oTempClass = New PS_CO923
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP560") Then '장비별 계획금액 등록(2015.10.21 추가)
            Set oTempClass = New PS_PP560
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP561") Then '장비별 예상금액(월별) 등록(2015.10.23 추가)
            Set oTempClass = New PS_PP561
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP562") Then '장비별 목표금액 수정관리(2017.03.30), 장비별 생산목표금액 분할(2016.01.26 추가)에서 변경
            Set oTempClass = New PS_PP562
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP563") Then '장비별 견적금액 수정관리(2017.03.30)
            Set oTempClass = New PS_PP563
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP564") Then '장비별 수주금액 수정관리(2017.03.30)
            Set oTempClass = New PS_PP564
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP565") Then '장비별 계획금액 등록(2015.10.22 추가)
            Set oTempClass = New PS_PP565
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD955") Then '판매마감 매출현황(2016.04.21 추가)
            Set oTempClass = New PS_SD955
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP567") Then '장비이익률관리 등록 현황 조회(2016.05.31 추가)
            Set oTempClass = New PS_PP567
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_FX030") Then '고정자산 현황 및 상태 등록(2016.07.21)
            Set oTempClass = New PS_FX030
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP570") Then '작번 비용 이관(2016.08.02)
            Set oTempClass = New PS_PP570
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM170") Then '현장 불합리 개선(2016.08.24)
            Set oTempClass = New PS_QM170
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD061") Then '기성매입등록및조회(2017.03.16)
            Set oTempClass = New PS_SD061
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_GA032") Then '월별 계획 등록(관리담당비용계획)(2017.04.26)
            Set oTempClass = New PS_GA032
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_GA033") Then '월별 계획VS실적 조회(관리담당비용계획)(2017.04.28)
            Set oTempClass = New PS_GA033
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_MM310") Then '구매품의 주문서 메일 발송(2017.05.22)
            Set oTempClass = New PS_MM310
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_GA059") Then '종합건강진단 희망병원 등록 및 조회(2017.05.25)
            Set oTempClass = New PS_GA059
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP250") Then '실패비용 관리(기계)(2017.06.09)
            Set oTempClass = New PS_PP250
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP568") Then '장비별 공정진행률 등록(2017.06.15)
            Set oTempClass = New PS_PP568
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP921") Then '설계비 배부 결과 조회(기계)(2017.06.26)
            Set oTempClass = New PS_PP921
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_GA500") Then '매출잔량조회(기계)(2017.06.30)
            Set oTempClass = New PS_GA500
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP991") Then '공수잔량조회(기계)(2017.08.17)
            Set oTempClass = New PS_PP991
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_GA501") Then '월별 장비/공구 자체/외주 매출 비율(기계)(2017.08.28)
            Set oTempClass = New PS_GA501
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD062") Then '월별기성매출현황조회(전체)(2017.09.04)
            Set oTempClass = New PS_SD062
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_QM092") Then '실패비용대상조회(2017.11.06)
            Set oTempClass = New PS_QM092
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_GA502") Then '월별 가동률 조회(2017.11.16)
            Set oTempClass = New PS_GA502
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_CO001") Then '결산마감관리(2017.11.21)
            Set oTempClass = New PS_CO001
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP251") Then '월별 가동률 관리(2018.01.18)
            Set oTempClass = New PS_PP251
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_SD956") Then '작번 이력 관리(2018.03.30)
            Set oTempClass = New PS_SD956
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_PP982") Then '반별가동률 조회(2018.05.16)
            Set oTempClass = New PS_PP982
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_CO660") Then '기간비용등록(2018.06.18)
            Set oTempClass = New PS_CO660
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_CO670") Then '기간비용분개등록(2018.06.21)
            Set oTempClass = New PS_CO670
            Call oTempClass.LoadForm
        ElseIf (pval.MenuUID = "PS_CO665") Then '기간비용현황(연간)(2018.06.27)
            Set oTempClass = New PS_CO665
            Call oTempClass.LoadForm
        End If

    End If
    Exit Sub
CreateUserForm_Error:
    Sbo_Application.SetStatusBarMessage "CreateUserForm_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub CreateSystemForm(ByRef pval As SAPbouiCOM.ItemEvent)
On Error GoTo CreateSystemForm_Error
    Dim oTempClass          As Variant
    If (pval.BeforeAction = True) Then
        If (pval.EventType = et_FORM_LOAD) Then
            '//판매관리
            If (pval.FormTypeEx = "149") Then '//판매견적
                Set oTempClass = New S149
                oTempClass.LoadForm (pval.FormUID) '//폼설정
                AddForms oTempClass, pval.FormUID '//폼추가
            ElseIf (pval.FormTypeEx = "139") Then '//판매오더
                Set oTempClass = New S139
                oTempClass.LoadForm (pval.FormUID)
                AddForms oTempClass, pval.FormUID
            ElseIf (pval.FormTypeEx = "140") Then '//납품
                Set oTempClass = New S140
                oTempClass.LoadForm (pval.FormUID)
                AddForms oTempClass, pval.FormUID
            ElseIf (pval.FormTypeEx = "180") Then '//판매반품
                Set oTempClass = New S180
                oTempClass.LoadForm (pval.FormUID)
                AddForms oTempClass, pval.FormUID
            ElseIf (pval.FormTypeEx = "133") Then '//AR송장
                Set oTempClass = New S133
                oTempClass.LoadForm (pval.FormUID)
                AddForms oTempClass, pval.FormUID
            ElseIf (pval.FormTypeEx = "179") Then '//AR대변메모
                Set oTempClass = New S179
                oTempClass.LoadForm (pval.FormUID)
                AddForms oTempClass, pval.FormUID
            ElseIf (pval.FormTypeEx = "60091") Then '//AR대변메모
                Set oTempClass = New S60091
                oTempClass.LoadForm (pval.FormUID)
                AddForms oTempClass, pval.FormUID
            '//구매관리
            ElseIf (pval.FormTypeEx = "142") Then '//구매오더
                Set oTempClass = New S142
                oTempClass.LoadForm (pval.FormUID)
                AddForms oTempClass, pval.FormUID
            ElseIf (pval.FormTypeEx = "143") Then '//입고PO
                Set oTempClass = New S143
                oTempClass.LoadForm (pval.FormUID)
                AddForms oTempClass, pval.FormUID
            ElseIf (pval.FormTypeEx = "182") Then '//구매반품
                Set oTempClass = New S182
                oTempClass.LoadForm (pval.FormUID)
                AddForms oTempClass, pval.FormUID
            ElseIf (pval.FormTypeEx = "141") Then '//AP송장
                Set oTempClass = New S141
                oTempClass.LoadForm (pval.FormUID)
                AddForms oTempClass, pval.FormUID
            ElseIf (pval.FormTypeEx = "181") Then '//AP대변메모
                Set oTempClass = New S181
                oTempClass.LoadForm (pval.FormUID)
                AddForms oTempClass, pval.FormUID
            ElseIf (pval.FormTypeEx = "60092") Then '//AP예약송장
                Set oTempClass = New S60092
                oTempClass.LoadForm (pval.FormUID)
                AddForms oTempClass, pval.FormUID
            '//회계관리
            ElseIf (pval.FormTypeEx = "804") Then '//계정과목
                Set oTempClass = New S804
                oTempClass.LoadForm (pval.FormUID)
                AddForms oTempClass, pval.FormUID
            ElseIf (pval.FormTypeEx = "392") Then '//분개
                Set oTempClass = New S392
                oTempClass.LoadForm (pval.FormUID)
                AddForms oTempClass, pval.FormUID
            ElseIf (pval.FormTypeEx = "393") Then '//분개장문서
                Set oTempClass = New S393
                oTempClass.LoadForm (pval.FormUID)
                AddForms oTempClass, pval.FormUID
            ElseIf (pval.FormTypeEx = "146") Then '//입금-지급수단
                Set oTempClass = New S146
                oTempClass.LoadForm (pval.FormUID)
                AddForms oTempClass, pval.FormUID
            ElseIf (pval.FormTypeEx = "196") Then '//지급-지급수단
                Set oTempClass = New S196
                oTempClass.LoadForm (pval.FormUID)
                AddForms oTempClass, pval.FormUID
            ElseIf (pval.FormTypeEx = "170") Then '//입금
                Set oTempClass = New S170
                oTempClass.LoadForm (pval.FormUID)
                AddForms oTempClass, pval.FormUID
            ElseIf (pval.FormTypeEx = "426") Then '//지급
                Set oTempClass = New S426
                oTempClass.LoadForm (pval.FormUID)
                AddForms oTempClass, pval.FormUID
            ElseIf (pval.FormTypeEx = "60051") Then '//어음관리
                Set oTempClass = New S60051
                oTempClass.LoadForm (pval.FormUID)
                AddForms oTempClass, pval.FormUID
            '//기타
            ElseIf (pval.FormTypeEx = "41") Then '//배치번호추가
                Set oTempClass = New S41
                oTempClass.LoadForm (pval.FormUID)
                AddForms oTempClass, pval.FormUID
            ElseIf (pval.FormTypeEx = "42") Then '//배치번호선택
                Set oTempClass = New S42
                oTempClass.LoadForm (pval.FormUID)
                AddForms oTempClass, pval.FormUID
            '//재고거래
            ElseIf (pval.FormTypeEx = "720") Then '//출고
                Set oTempClass = New S720
                oTempClass.LoadForm (pval.FormUID)
                AddForms oTempClass, pval.FormUID
            ElseIf (pval.FormTypeEx = "721") Then '//입고
                Set oTempClass = New S721
                oTempClass.LoadForm (pval.FormUID)
                AddForms oTempClass, pval.FormUID
            
            ElseIf (pval.FormTypeEx = "150") Then '품목마스터
                Set oTempClass = New S150
                oTempClass.LoadForm (pval.FormUID)
                AddForms oTempClass, pval.FormUID
            
            End If
        End If
    End If
    Exit Sub
CreateSystemForm_Error:
    Sbo_Application.SetStatusBarMessage "CreateSystemForm_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
