VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY985"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//  SAP MANAGE UI API 2004 SDK Sample
'//****************************************************************************
'//  File           : PH_PY985.cls
'//  Module         : 인사관리>정산관리
'//  Desc           : 의료비지급명세서자료 전산매체수록
'//  FormType       :
'//  Create Date    : 2014.01.28
'//  Modified Date  : 2016.01.10
'//  Creator        : NGY
'//  Modifier       :
'//  Copyright  (c) Poongsan Holdings
'//****************************************************************************

Option Explicit

Public oFormUniqueID    As String
Public oForm            As SAPbouiCOM.Form
Private sRecordset      As SAPbobsCOM.Recordset
Private oMat1           As SAPbouiCOM.Matrix
Private Last_Item       As String                     '클래스에서 선택한 마지막 아이템 Uid값

Private CLTCOD          As String
Private yyyy            As String
Private HtaxID          As String
Private DocDate         As String
Private oFilePath       As String

Private FILNAM      As String * 30    '파  일  명
Private MaxRow      As Long
Private BUSCNT      As Integer  '/ B레코드일련번호
Private BUSTOT      As Integer  '/ B레코드총갯수

Private NEWCNT      As Integer
Private OLDCNT      As Integer
Private C_SAUP      As String
Private C_YYYY      As String
Private C_SABUN     As String
Private E_BUYCNT    As String
Private C_BUYCNT    As String


'2013년기준 250 BYTE
'2015년기준 251 BYTE
Private Type A_record
        A001   As String * 1     '레코드구분 'A'
        A002   As String * 2     '자료구분   '26'
        A003   As String * 3     '세무서
        A004   As String * 6     '일련번호
        A005   As String * 8     '제출년월일
        A006   As String * 10    '사업자번호
        A007   As String * 20    '홈텍스ID
        A008   As String * 4     '세무프로그램코드
        A009   As String * 10    '사업자번호
        A010   As String * 40    '법인명(상호)
        A011   As String * 13    '소득자의주민등록번호
        A012   As String * 1     '내,외국인
        A013   As String * 30    '성명
        A014   As String * 10    '지급처사업자등록번호
        A015   As String * 40    '지급처상호
        A016   As String * 1     '의료증빙코드
        A017   As String * 5     '건수
        A018   As String * 11    '지급금액
        A019   As String * 1     '난임시술비
        A020   As String * 13    '주민등록번호
        A021   As String * 1     '내,외국인코드
        A022   As String * 1     '본인등해당여부
        A023   As String * 1     '제출대상기간코드
        A024   As String * 19    '공란
End Type
Dim A_rec As A_record

'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error
    Dim i           As Long
    Dim oXmlDoc             As New MSXML2.DOMDocument

    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY985.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = _
        oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '//여러개의 메트릭스가 틀경우에 층계모양처럼 로드 되도록 만든 모양
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetTotalFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetTotalFormsCount * 10)

    Sbo_Application.LoadBatchActions oXmlDoc.xml
    
    oFormUniqueID = "PH_PY985_" & GetTotalFormsCount
    
    '폼 할당
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)
    
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '컬렉션에 폼을 담는다   **컬렉션이란 개체를 담아 놓는 배열로서 여기서는 활성화되어져 있는 폼을 담고 있다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    AddForms Me, oFormUniqueID, "PH_PY985"
    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
    
    oForm.Freeze True
    CreateItems
    oForm.Freeze False
    
    oForm.EnableMenu ("1281"), False '/ 찾기
    oForm.EnableMenu ("1282"), True  '/ 추가
    oForm.EnableMenu ("1284"), False '/ 취소
    oForm.EnableMenu ("1293"), False '/ 행삭제
    oForm.Update
    oForm.Visible = True

    Set oXmlDoc = Nothing
    Exit Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
LoadForm_Error:
    Set oXmlDoc = Nothing
    Sbo_Application.StatusBar.SetText "Form_Load Error:" & Err.Description, bmt_Short, smt_Error
    If (oForm Is Nothing) = False Then
        oForm.Freeze False
        Set oForm = Nothing
    End If
End Sub
'*******************************************************************
'// ItemEventHander
'*******************************************************************
Public Sub Raise_FormItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
    
    Dim sQry        As String
    Dim i           As Long
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns     As SAPbouiCOM.Columns
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo Raise_FormItemEvent_Error
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Select Case pval.EventType
'et_ITEM_PRESSED''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_ITEM_PRESSED
            If pval.BeforeAction Then
'
                If pval.ItemUID = "1" Then
                    If oForm.Mode = fm_ADD_MODE Then
                        If HeaderSpaceLineDel = False Then
                                BubbleEvent = False
                                Exit Sub
                        End If
                    End If
                End If
                If pval.ItemUID = "Btn01" Then
                   If File_Create = False Then
                       BubbleEvent = False
                       Exit Sub
                   Else
                       BubbleEvent = False
                       oForm.Mode = fm_OK_MODE
                   End If
                
                End If
            Else
            End If
            
        Case et_COMBO_SELECT:
            If pval.BeforeAction = True Then
                
            ElseIf pval.BeforeAction = False Then
                If pval.ItemChanged = True Then
                    Select Case pval.ItemUID
                           '//사업장이 바뀌면
                           Case "CLTCOD"
                                sQry = "SELECT U_HomeTId, U_ChgDpt, U_ChgName, U_ChgTel  FROM [@PH_PY005A] WHERE U_CLTCode = '" & Trim(oForm.Items("CLTCOD").Specific.VALUE) & "'"
                                oRecordSet.DoQuery sQry
                                oForm.Items("HtaxID").Specific.String = Trim(oRecordSet.Fields("U_HomeTId").VALUE)
                    End Select
                End If
            End If
'et_VALIDATE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_VALIDATE
    
'et_CLICK''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_CLICK
    
'et_KEY_DOWN''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_KEY_DOWN
    
'et_GOT_FOCUS''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_GOT_FOCUS
    
'et_FORM_UNLOAD''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_FORM_UNLOAD
            'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
            '컬렉션에서 삭제및 모든 메모리 제거
            'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
            If pval.BeforeAction = False Then
               RemoveForms oFormUniqueID
                Set oForm = Nothing
                Set oMat1 = Nothing
            End If
    End Select
    
    Exit Sub
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_FormItemEvent_Error:
    Sbo_Application.StatusBar.SetText "Raise_FormItemEvent_Error:" & Space$(10) & Err.Description, bmt_Short, smt_Error
End Sub

'*******************************************************************
'// MenuEventHander
'*******************************************************************
Public Sub Raise_FormMenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
    
    If pval.BeforeAction = True Then
        Exit Sub
    End If
    
    Select Case pval.MenuUID
        Case "1287" '/ 복제
        Case "1281", "1282":
            oForm.Items("JsnYear").CLICK ct_Regular
        Case "1288" To "1291":
        Case "1293"
    End Select
    Exit Sub
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

Private Sub CreateItems()
On Error GoTo Error_Message
    Dim oCombo          As SAPbouiCOM.ComboBox
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String

    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Call oForm.DataSources.UserDataSources.Add("CLTCOD", dt_SHORT_TEXT, 10)
    Set oCombo = oForm.Items("CLTCOD").Specific
    oCombo.DataBind.SetBound True, "", "CLTCOD"
    oForm.Items("CLTCOD").DisplayDesc = True
    '// 접속자에 따른 권한별 사업장 콤보박스세팅
    Call CLTCOD_Select(oForm, "CLTCOD")
    
    oForm.Items("YYYY").Specific.String = Format(Now, "YYYY") - 1    '년도 기본년도에서 - 1
    
    Call oForm.DataSources.UserDataSources.Add("DocDate", dt_DATE, 10)     '제출일자
    oForm.Items("DocDate").Specific.DataBind.SetBound True, "", "DocDate"
    
    Set oRecordSet = Nothing
    Exit Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Error_Message:
    Set oRecordSet = Nothing
    Sbo_Application.StatusBar.SetText "CreateItems 실행 중 오류가 발생했습니다." & Space$(10) & Err.Description, bmt_Short, smt_Error
End Sub

Private Function File_Create() As Boolean
On Error GoTo Error_Message
    Dim ErrNum          As Integer
    Dim oStr            As String
    Dim sQry            As String
    
    Set sRecordset = oCompany.GetBusinessObject(BoRecordset)
    
    '화면변수를 전역변수로 MOVE
    CLTCOD = Trim(oForm.Items("CLTCOD").Specific.VALUE)
    yyyy = Trim(oForm.Items("YYYY").Specific.VALUE)
    HtaxID = Trim(oForm.Items("HtaxID").Specific.VALUE)
    DocDate = Trim(oForm.Items("DocDate").Specific.VALUE)
    
    ErrNum = 0
    
    '/ Question
    If Sbo_Application.MessageBox("의료비 신고파일을 생성하시겠습니까?", 2, "&Yes!", "&No") = 2 Then
        ErrNum = 1
        GoTo Error_Message
    End If
    
    '/ A RECORD 처리
    If File_Create_A_record = False Then
        ErrNum = 2
        GoTo Error_Message
    End If
    
    Close #1
    
    Sbo_Application.StatusBar.SetText "전산매체수록이 정상적으로 완료되었습니다.", bmt_Short, smt_Success
    File_Create = True
    Set sRecordset = Nothing
    Exit Function
'///////////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set sRecordset = Nothing
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "취소하였습니다.", bmt_Short, smt_Success
    ElseIf ErrNum = 2 Then
        Sbo_Application.StatusBar.SetText "A레코드 생성 실패.", bmt_Short, smt_Error
    Else
        Sbo_Application.StatusBar.SetText "File_Create 실행 중 오류가 발생했습니다." & Space$(10) & Err.Description, bmt_Short, smt_Error
    End If
    File_Create = False
End Function
Private Function File_Create_A_record() As Boolean
On Error GoTo Error_Message
    Dim ErrNum          As Integer
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim PRTDAT          As String
    Dim saup            As String
    Dim CheckA          As String
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    CheckA = False  '/체크필요유무
    ErrNum = 0
    
    '/ A_RECORE QUERY
    sQry = "EXEC PH_PY985_A '" & CLTCOD & "', '" & yyyy & "', '" & HtaxID & "', '" & DocDate & "'"
    oRecordSet.DoQuery sQry
        
    If oRecordSet.RecordCount = 0 Then
        ErrNum = 1
        GoTo Error_Message
    Else
        'PATH및 파일이름 만들기
        saup = oRecordSet.Fields("A009").VALUE  '사업자번호
        oFilePath = "C:\BANK\CA" + Mid$(saup, 1, 7) + "." + Mid$(saup, 8, 3)
        Close #1
        Open oFilePath For Output As #1
        
        Dim ProgressBar01 As SAPbouiCOM.ProgressBar
        Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("작성시작!", oRecordSet.RecordCount, False)
    
        NEWCNT = 0
        
        Do Until oRecordSet.EOF
        
            NEWCNT = NEWCNT + 1     '/ 일련번호
            
            'A RECORD MOVE
            
            A_rec.A001 = oRecordSet.Fields("A001").VALUE
            A_rec.A002 = oRecordSet.Fields("A002").VALUE
            A_rec.A003 = oRecordSet.Fields("A003").VALUE
            A_rec.A004 = Format$(NEWCNT, String$(Len(A_rec.A004), "0"))  '/ 일련번호
            A_rec.A005 = oRecordSet.Fields("A005").VALUE
            A_rec.A006 = oRecordSet.Fields("A006").VALUE
            A_rec.A007 = oRecordSet.Fields("A007").VALUE
            A_rec.A008 = oRecordSet.Fields("A008").VALUE
            A_rec.A009 = oRecordSet.Fields("A009").VALUE
            A_rec.A010 = oRecordSet.Fields("A010").VALUE
            A_rec.A011 = oRecordSet.Fields("A011").VALUE
            A_rec.A012 = oRecordSet.Fields("A012").VALUE
            A_rec.A013 = oRecordSet.Fields("A013").VALUE
            A_rec.A014 = oRecordSet.Fields("A014").VALUE
            A_rec.A015 = oRecordSet.Fields("A015").VALUE
            A_rec.A016 = oRecordSet.Fields("A016").VALUE
            A_rec.A017 = Format$(oRecordSet.Fields("A017").VALUE, String$(Len(A_rec.A017), "0"))
            A_rec.A018 = Format$(oRecordSet.Fields("A018").VALUE, String$(Len(A_rec.A018), "0"))
            A_rec.A019 = oRecordSet.Fields("A019").VALUE
            A_rec.A020 = oRecordSet.Fields("A020").VALUE
            A_rec.A021 = oRecordSet.Fields("A021").VALUE
            A_rec.A022 = oRecordSet.Fields("A022").VALUE
            A_rec.A023 = oRecordSet.Fields("A023").VALUE
            A_rec.A024 = oRecordSet.Fields("A024").VALUE
            
            
            
            Print #1, MDC_SetMod.sStr(A_rec.A001) & MDC_SetMod.sStr(A_rec.A002) & MDC_SetMod.sStr(A_rec.A003) & MDC_SetMod.sStr(A_rec.A004) & _
                      MDC_SetMod.sStr(A_rec.A005) & MDC_SetMod.sStr(A_rec.A006) & MDC_SetMod.sStr(A_rec.A007) & MDC_SetMod.sStr(A_rec.A008) & _
                      MDC_SetMod.sStr(A_rec.A009) & MDC_SetMod.sStr(A_rec.A010) & MDC_SetMod.sStr(A_rec.A011) & MDC_SetMod.sStr(A_rec.A012) & _
                      MDC_SetMod.sStr(A_rec.A013) & MDC_SetMod.sStr(A_rec.A014) & MDC_SetMod.sStr(A_rec.A015) & MDC_SetMod.sStr(A_rec.A016) & _
                      MDC_SetMod.sStr(A_rec.A017) & MDC_SetMod.sStr(A_rec.A018) & MDC_SetMod.sStr(A_rec.A019) & MDC_SetMod.sStr(A_rec.A020) & _
                      MDC_SetMod.sStr(A_rec.A021) & MDC_SetMod.sStr(A_rec.A022) & MDC_SetMod.sStr(A_rec.A023) & MDC_SetMod.sStr(A_rec.A024)
                      
            oRecordSet.MoveNext
            
            ProgressBar01.VALUE = ProgressBar01.VALUE + 1
            ProgressBar01.Text = ProgressBar01.VALUE & "/" & oRecordSet.RecordCount & "건 작성중........!"
        
            
        Loop
    End If
    
    If CheckA = False Then
        File_Create_A_record = True
    Else
        File_Create_A_record = False
    End If
    Set oRecordSet = Nothing
    Exit Function
'///////////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set oRecordSet = Nothing
   
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "의료비자료가 존재하지 않습니다. 등록하여 주십시오.", bmt_Short, smt_Error
    Else
        Call Matrix_AddRow("A레코드오류: " & Err.Description, False, True)
    End If
    
    File_Create_A_record = False
    
End Function


Private Sub Matrix_AddRow(ByVal MatrixMsg As String, Optional Insert_YN As Boolean, Optional MatrixErr As Boolean)
    If MatrixErr = True Then
        oForm.DataSources.UserDataSources("Col0").VALUE = "??"
    Else
        oForm.DataSources.UserDataSources("Col0").VALUE = ""
    End If
    oForm.DataSources.UserDataSources("Col1").VALUE = MatrixMsg
    If Insert_YN = True Then
        oMat1.AddRow
        MaxRow = MaxRow + 1
    End If
    oMat1.SetLineData MaxRow
End Sub


'화면변수 CHECK
Private Function HeaderSpaceLineDel() As Boolean
On Error GoTo HeaderSpaceLineDel
    Dim ErrNum      As Integer
    
    ErrNum = 0
    '/ 필수Check
    If oForm.Items("HtaxID").Specific.VALUE = "" Then
        ErrNum = 1
        GoTo HeaderSpaceLineDel
    ElseIf oForm.Items("DocDate").Specific.VALUE = "" Then
        ErrNum = 2
        GoTo HeaderSpaceLineDel
    End If
    
    HeaderSpaceLineDel = True
    Exit Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
HeaderSpaceLineDel:
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "홈텍스ID(5자리이상)를 확인하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 2 Then
        Sbo_Application.StatusBar.SetText "제출일자는 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
    Else
        Sbo_Application.StatusBar.SetText "HeaderSpaceLineDel 실행 중 오류가 발생했습니다." & Space$(10) & Err.Description, bmt_Short, smt_Error
    End If
    
    HeaderSpaceLineDel = False
End Function


