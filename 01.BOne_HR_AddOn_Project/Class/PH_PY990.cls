VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY990"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//  SAP MANAGE UI API 2004 SDK Sample
'//****************************************************************************
'//  File           : PH_PY990.cls
'//  Module         : 인사관리>정산관리
'//  Desc           : 기부금지급명세서자료 전산매체수록
'//  FormType       :
'//  Create Date    : 2014.02.03
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
Private TeamName        As String
Private Dname           As String
Private Dtel            As String
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


'2013년기준 180 BYTE
'2016년기준 190 BYTE
Private Type A_record
        A001   As String * 1     '레코드구분
        A002   As String * 2     '자료구분
        A003   As String * 3     '세무서
        A004   As String * 8     '제출일자
        A005   As String * 1     '제출자구분 (1;세무대리인, 2;법인, 3;개인)
        A006   As String * 6     '세무대리인
        A007   As String * 20    '홈텍스ID
        A008   As String * 4     '세무프로그램코드
        A009   As String * 10    '사업자번호
        A010   As String * 40    '법인명(상호)
        A011   As String * 30    '담당자부서
        A012   As String * 30    '담당자성명
        A013   As String * 15    '담당자전화번호
        A014   As String * 5     '신고의무자수
        A015   As String * 3     '한글코드종류
        A016   As String * 12    '공란
End Type
Dim A_rec As A_record


Private Type B_record
        B001   As String * 1     '레코드구분
        B002   As String * 2     '자료구분
        B003   As String * 3     '세무서
        B004   As String * 6     '일련번호
        B005   As String * 10    '사업자번호
        B006   As String * 40    '법인명(상호)
        B007   As String * 7     'C레코드수
        B008   As String * 7     'D레코드수
        B009   As String * 13    '기부금액합계
        B010   As String * 13    '공제대상금액합계
        B011   As String * 1     '제출대상기간코드
        B012   As String * 87    '공란
End Type
Dim B_rec As B_record


Private Type C_record
        C001   As String * 1     '레코드구분
        C002   As String * 2     '자료구분
        C003   As String * 3     '세무서
        C004   As String * 6     '일련번호
        C005   As String * 10    '사업자번호
        C006   As String * 13    '소득자주민등록번호
        C007   As String * 1     '내,외국인코드
        C008   As String * 30    '성명
        C009   As String * 2     '유형코드
        C010   As String * 4     '기부년도
        C011   As String * 13    '기부금액
        C012   As String * 13    '전년까지공제된금액
        C013   As String * 13    '공제대상금액
        C014   As String * 13    '해당년도공제금액 필요경비 '0'  2016
        C015   As String * 13    '해당년도공제금액세액(소득)공제금액  2016
        C016   As String * 13    '해당년도에공제받지못한금액_소멸금액
        C017   As String * 13    '해당년도에공제받지못한금액_이월금액
        C018   As String * 5     '기부조정명세일련번호
        C019   As String * 22    '공란
End Type
Dim C_rec As C_record

Private Type D_Record
        D001   As String * 1    '레코드구분
        D002   As String * 2    '자료구분
        D003   As String * 3    '세무서
        D004   As String * 6    '일련번호
        D005   As String * 10   '사업자등록번호
        D006   As String * 13   '주민등록번호
        D007   As String * 2    '유형코드
        D008   As String * 13   '기부처-사업자등록번호
        D009   As String * 30   '기부처-법인명(상호)
        D010   As String * 1    '관계
        D011   As String * 1    '내,외국인코드
        D012   As String * 20   '성명
        D013   As String * 13   '주민등록번호
        D014   As String * 5    '건수
        D015   As String * 13   '금액
        D016   As String * 13   '공제대상기부금액  2016
        D017   As String * 13   '기부장려금신청금액  2016
        D018   As String * 5    '해당연도기부명세일련번호
        D019   As String * 26   '공란
End Type
Dim D_rec As D_Record

'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error
    Dim i           As Long
    Dim oXmlDoc             As New MSXML2.DOMDocument

    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY990.srf")
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
    
    oFormUniqueID = "PH_PY990_" & GetTotalFormsCount
    
    '폼 할당
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)
    
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '컬렉션에 폼을 담는다   **컬렉션이란 개체를 담아 놓는 배열로서 여기서는 활성화되어져 있는 폼을 담고 있다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    AddForms Me, oFormUniqueID, "PH_PY990"
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
'                If pval.ItemUID = "CBtn1" Then   '/ ChooseBtn사원리스트
'                    oForm.Items("MSTCOD").CLICK ct_Regular
'                    Sbo_Application.ActivateMenuItem ("7425")
'                    BubbleEvent = False
'                Else
                 
                 If pval.ItemUID = "Btn01" Then
                    If HeaderSpaceLineDel = False Then
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
                                oForm.Items("TeamName").Specific.String = Trim(oRecordSet.Fields("U_ChgDpt").VALUE)
                                oForm.Items("Dname").Specific.String = Trim(oRecordSet.Fields("U_ChgName").VALUE)
                                oForm.Items("Dtel").Specific.String = Trim(oRecordSet.Fields("U_ChgTel").VALUE)
                           
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
    TeamName = Trim(oForm.Items("TeamName").Specific.VALUE)
    Dname = Trim(oForm.Items("Dname").Specific.VALUE)
    Dtel = Trim(oForm.Items("Dtel").Specific.VALUE)
    DocDate = Trim(oForm.Items("DocDate").Specific.VALUE)
    
    ErrNum = 0
    
    '/ Question
    If Sbo_Application.MessageBox("전산매체신고 파일을 생성하시겠습니까?", 2, "&Yes!", "&No") = 2 Then
        ErrNum = 1
        GoTo Error_Message
    End If
    
    '/ A RECORD 처리
    If File_Create_A_record = False Then
        ErrNum = 2
        GoTo Error_Message
    End If
    
    '/ B RECORD 처리
    If File_Create_B_record = False Then
        ErrNum = 3
        GoTo Error_Message
    End If
    
    '/ C RECORD 처리  D 처리
    If File_Create_C_record = False Then
        ErrNum = 4
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
        Sbo_Application.StatusBar.SetText "A레코드(제출자 레코드) 생성 실패.", bmt_Short, smt_Error
    ElseIf ErrNum = 3 Then
        Sbo_Application.StatusBar.SetText "B레코드(원천징수의무자별 집계 레코드) 생성 실패.", bmt_Short, smt_Error
    ElseIf ErrNum = 4 Then
        Sbo_Application.StatusBar.SetText "C레코드(기부금 조정명세 레코드) 생성 실패.", bmt_Short, smt_Error
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
    sQry = "EXEC PH_PY990_A '" & CLTCOD & "', '" & HtaxID & "', '" & TeamName & "', '" & Dname & "', '" & Dtel & "', '" & DocDate & "'"
    oRecordSet.DoQuery sQry
    
    If oRecordSet.RecordCount = 0 Then
        ErrNum = 1
        GoTo Error_Message
    Else
        ' PATH및 파일이름 만들기
        saup = oRecordSet.Fields("A009").VALUE  '사업자번호
        oFilePath = "C:\BANK\H" + Mid$(saup, 1, 7) + "." + Mid$(saup, 8, 3)
        
        
        'A RECORD MOVE
        
        A_rec.A001 = oRecordSet.Fields("A001").VALUE
        A_rec.A002 = oRecordSet.Fields("A002").VALUE
        A_rec.A003 = oRecordSet.Fields("A003").VALUE
        A_rec.A004 = oRecordSet.Fields("A004").VALUE
        A_rec.A005 = oRecordSet.Fields("A005").VALUE
        A_rec.A006 = oRecordSet.Fields("A006").VALUE
        A_rec.A007 = oRecordSet.Fields("A007").VALUE
        A_rec.A008 = oRecordSet.Fields("A008").VALUE
        A_rec.A009 = oRecordSet.Fields("A009").VALUE
        A_rec.A010 = oRecordSet.Fields("A010").VALUE
        A_rec.A011 = oRecordSet.Fields("A011").VALUE
        A_rec.A012 = oRecordSet.Fields("A012").VALUE
        A_rec.A013 = oRecordSet.Fields("A013").VALUE
        
        A_rec.A014 = Format$(oRecordSet.Fields("A014").VALUE, String$(Len(A_rec.A014), "0"))
        A_rec.A015 = oRecordSet.Fields("A015").VALUE
        A_rec.A016 = oRecordSet.Fields("A016").VALUE
        
        Close #1
        Open oFilePath For Output As #1
        Print #1, MDC_SetMod.sStr(A_rec.A001) & MDC_SetMod.sStr(A_rec.A002) & MDC_SetMod.sStr(A_rec.A003) & MDC_SetMod.sStr(A_rec.A004) & _
                  MDC_SetMod.sStr(A_rec.A005) & MDC_SetMod.sStr(A_rec.A006) & MDC_SetMod.sStr(A_rec.A007) & MDC_SetMod.sStr(A_rec.A008) & _
                  MDC_SetMod.sStr(A_rec.A009) & MDC_SetMod.sStr(A_rec.A010) & MDC_SetMod.sStr(A_rec.A011) & MDC_SetMod.sStr(A_rec.A012) & _
                  MDC_SetMod.sStr(A_rec.A013) & MDC_SetMod.sStr(A_rec.A014) & MDC_SetMod.sStr(A_rec.A015) & MDC_SetMod.sStr(A_rec.A016)
                  
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
        Sbo_Application.StatusBar.SetText "귀속년도의 자사정보(A RECORD)가 존재하지 않습니다. 등록하여 주십시오.", bmt_Short, smt_Error
    Else
        Call Matrix_AddRow("A레코드오류: " & Err.Description, False, True)
    End If
    
    File_Create_A_record = False
    
End Function

Private Function File_Create_B_record() As Integer
On Error GoTo Error_Message
    Dim ErrNum          As Integer
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim CheckB          As String
    
    CheckB = False  '/체크필요유무
    ErrNum = 0
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '/ B_RECORE QUERY
    sQry = "EXEC PH_PY990_B '" & CLTCOD & "', '" & yyyy & "'"
    oRecordSet.DoQuery sQry
    
    If oRecordSet.RecordCount = 0 Then
        ErrNum = 1
        GoTo Error_Message
    Else
        'B RECORD MOVE
        
        B_rec.B001 = oRecordSet.Fields("B001").VALUE
        B_rec.B002 = oRecordSet.Fields("B002").VALUE
        B_rec.B003 = oRecordSet.Fields("B003").VALUE
        B_rec.B004 = oRecordSet.Fields("B004").VALUE
        B_rec.B005 = oRecordSet.Fields("B005").VALUE
        B_rec.B006 = oRecordSet.Fields("B006").VALUE
        B_rec.B007 = Format$(oRecordSet.Fields("B007").VALUE, String$(Len(B_rec.B007), "0"))
        B_rec.B008 = Format$(oRecordSet.Fields("B008").VALUE, String$(Len(B_rec.B008), "0"))
        B_rec.B009 = Format$(oRecordSet.Fields("B009").VALUE, String$(Len(B_rec.B009), "0"))
        B_rec.B010 = Format$(oRecordSet.Fields("B010").VALUE, String$(Len(B_rec.B010), "0"))
        B_rec.B011 = oRecordSet.Fields("B011").VALUE
        B_rec.B012 = oRecordSet.Fields("B012").VALUE
        
        Print #1, MDC_SetMod.sStr(B_rec.B001) & MDC_SetMod.sStr(B_rec.B002) & MDC_SetMod.sStr(B_rec.B003) & MDC_SetMod.sStr(B_rec.B004) & _
                  MDC_SetMod.sStr(B_rec.B005) & MDC_SetMod.sStr(B_rec.B006) & MDC_SetMod.sStr(B_rec.B007) & MDC_SetMod.sStr(B_rec.B008) & _
                  MDC_SetMod.sStr(B_rec.B009) & MDC_SetMod.sStr(B_rec.B010) & MDC_SetMod.sStr(B_rec.B011) & MDC_SetMod.sStr(B_rec.B012)
                  
    End If
    
    If CheckB = False Then
        File_Create_B_record = True
    Else
        File_Create_B_record = False
    End If
    
    Set oRecordSet = Nothing
    Exit Function
'///////////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set oRecordSet = Nothing
   
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "B레코드가 존재하지 않습니다.", bmt_Short, smt_Error
        File_Create_B_record = 1
    Else
        Call Matrix_AddRow("B레코드오류: " & Err.Description, False)
        File_Create_B_record = 1
    End If
    
End Function

Private Function File_Create_C_record() As Boolean
On Error GoTo Error_Message
    Dim ErrNum      As Integer
    Dim oRecordSet  As SAPbobsCOM.Recordset
    Dim sQry        As String
    Dim CheckC      As String
    Dim PSABUN      As String
    Dim OLDBIG      As Double
    Dim PILTOT      As Double
    Dim SCount      As Integer
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    CheckC = False  '/체크필요유무
    ErrNum = 0
    
    '/ C_RECORE QUERY
    sQry = "EXEC PH_PY990_C '" & CLTCOD & "', '" & yyyy & "'"
    
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
        ErrNum = 1
        GoTo Error_Message
    End If
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("작성시작!", oRecordSet.RecordCount, False)
    
    NEWCNT = 1
    SCount = 0
    PSABUN = oRecordSet.Fields("sabun").VALUE
    
    Do Until oRecordSet.EOF
        
        C_SAUP = oRecordSet.Fields("saup").VALUE
        C_YYYY = oRecordSet.Fields("yyyy").VALUE
        C_SABUN = oRecordSet.Fields("sabun").VALUE
        
        'C RECORD MOVE
        
        SCount = SCount + 1
        
        C_rec.C001 = oRecordSet.Fields("C001").VALUE
        C_rec.C002 = oRecordSet.Fields("C002").VALUE
        C_rec.C003 = oRecordSet.Fields("C003").VALUE
        C_rec.C004 = Format$(NEWCNT, String$(Len(C_rec.C004), "0"))  '/ 일련번호
        C_rec.C005 = oRecordSet.Fields("C005").VALUE
        C_rec.C006 = oRecordSet.Fields("C006").VALUE
        C_rec.C007 = oRecordSet.Fields("C007").VALUE
        C_rec.C008 = oRecordSet.Fields("C008").VALUE
        C_rec.C009 = oRecordSet.Fields("C009").VALUE
        C_rec.C010 = oRecordSet.Fields("C010").VALUE
        
        C_rec.C011 = Format$(oRecordSet.Fields("C011").VALUE, String$(Len(C_rec.C011), "0"))
        C_rec.C012 = Format$(oRecordSet.Fields("C012").VALUE, String$(Len(C_rec.C012), "0"))
        C_rec.C013 = Format$(oRecordSet.Fields("C013").VALUE, String$(Len(C_rec.C013), "0"))
        C_rec.C014 = Format$(oRecordSet.Fields("C014").VALUE, String$(Len(C_rec.C014), "0"))
        C_rec.C015 = Format$(oRecordSet.Fields("C015").VALUE, String$(Len(C_rec.C015), "0"))
        C_rec.C016 = Format$(oRecordSet.Fields("C016").VALUE, String$(Len(C_rec.C016), "0"))
        C_rec.C017 = Format$(oRecordSet.Fields("C017").VALUE, String$(Len(C_rec.C017), "0"))
        C_rec.C018 = Format$(SCount, String$(Len(C_rec.C018), "0"))  '/ 일련번호
        C_rec.C019 = oRecordSet.Fields("C019").VALUE
        
        
        '예제
        'C_rec.PERNBR = Replace(oRecordSet.Fields("U_PERNBR").VALUE, "-", "")
        
        'OLDBIG = Val(oRecordSet.Fields("U_BIGWA1").VALUE) + Val(oRecordSet.Fields("U_BIGWA3").VALUE) + Val(oRecordSet.Fields("U_BIGWA5").VALUE) _
        '        + Val(oRecordSet.Fields("U_BIGWA6").VALUE) + Val(oRecordSet.Fields("U_BIGWU3").VALUE)
                              
        'C_rec.FILD02 = Format$(0, String$(Len(C_rec.FILD02), "0"))
        'C_rec.GAMFLD = String$(Len(C_rec.GAMFLD), "0")
        'C_rec.FILLER = Space$(Len(C_rec.FILLER))
        'C_rec.C022 = Format$(oRecordSet.Fields("C022").VALUE, , String$(Len(C_rec.C022), "0"))
        
      
        Print #1, MDC_SetMod.sStr(C_rec.C001) & MDC_SetMod.sStr(C_rec.C002) & MDC_SetMod.sStr(C_rec.C003) & MDC_SetMod.sStr(C_rec.C004) & MDC_SetMod.sStr(C_rec.C005) & MDC_SetMod.sStr(C_rec.C006) & MDC_SetMod.sStr(C_rec.C007) & MDC_SetMod.sStr(C_rec.C008) & MDC_SetMod.sStr(C_rec.C009) & MDC_SetMod.sStr(C_rec.C010) & _
                  MDC_SetMod.sStr(C_rec.C011) & MDC_SetMod.sStr(C_rec.C012) & MDC_SetMod.sStr(C_rec.C013) & MDC_SetMod.sStr(C_rec.C014) & MDC_SetMod.sStr(C_rec.C015) & MDC_SetMod.sStr(C_rec.C016) & MDC_SetMod.sStr(C_rec.C017) & MDC_SetMod.sStr(C_rec.C018) & MDC_SetMod.sStr(C_rec.C019)
                  
        
        oRecordSet.MoveNext
        
        
        ProgressBar01.VALUE = ProgressBar01.VALUE + 1
        ProgressBar01.Text = ProgressBar01.VALUE & "/" & oRecordSet.RecordCount & "건 작성중........!"
        
        
        If oRecordSet.EOF Then
           '/ D레코드
           If File_Create_D_record = False Then
              ErrNum = 2
              GoTo Error_Message
           End If

        ElseIf PSABUN <> oRecordSet.Fields("sabun").VALUE Then
           '/ D레코드
           If File_Create_D_record = False Then
              ErrNum = 2
              GoTo Error_Message
           End If

           NEWCNT = NEWCNT + 1     '/ 일련번호
           SCount = 0
           PSABUN = oRecordSet.Fields("sabun").VALUE

        End If
        
    Loop
        
        
    If CheckC = False Then
        File_Create_C_record = True
    Else
        File_Create_C_record = False
    End If
    Set oRecordSet = Nothing
    Exit Function
'///////////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set oRecordSet = Nothing
   
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "C레코드가 존재하지 않습니다. 등록하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 2 Then
        Sbo_Application.StatusBar.SetText "D레코드 생성 실패.", bmt_Short, smt_Error
    Else
        Call Matrix_AddRow("C레코드오류: " & Err.Description, False)
    End If
    File_Create_C_record = False
End Function

Private Function File_Create_D_record() As Boolean
On Error GoTo Error_Message
    Dim ErrNum      As Integer
    Dim oRecordSet  As SAPbobsCOM.Recordset
    Dim sQry        As String
    Dim CheckD      As String
    Dim DCount      As Integer
    
    CheckD = False  '/체크필요유무
    ErrNum = 0
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '/ D_RECORE QUERY
    sQry = "EXEC PH_PY990_D '" & C_SAUP & "', '" & C_YYYY & "', '" & C_SABUN & "'"
    
    oRecordSet.DoQuery sQry
    
    DCount = 0
    Do Until oRecordSet.EOF
    
        'D RECORD MOVE
        DCount = DCount + 1
        D_rec.D001 = oRecordSet.Fields("D001").VALUE
        D_rec.D002 = oRecordSet.Fields("D002").VALUE
        D_rec.D003 = oRecordSet.Fields("D003").VALUE
        D_rec.D004 = Format$(C_rec.C004, String$(Len(D_rec.D004), "0"))  '/ C레코드의 일련번호
        D_rec.D005 = oRecordSet.Fields("D005").VALUE
        D_rec.D006 = oRecordSet.Fields("D006").VALUE
        D_rec.D007 = oRecordSet.Fields("D007").VALUE
        D_rec.D008 = oRecordSet.Fields("D008").VALUE
        D_rec.D009 = oRecordSet.Fields("D009").VALUE
        D_rec.D010 = oRecordSet.Fields("D010").VALUE
        D_rec.D011 = oRecordSet.Fields("D011").VALUE
        D_rec.D012 = oRecordSet.Fields("D012").VALUE
        D_rec.D013 = oRecordSet.Fields("D013").VALUE
        D_rec.D014 = Format$(oRecordSet.Fields("D014").VALUE, String$(Len(D_rec.D014), "0"))
        D_rec.D015 = Format$(oRecordSet.Fields("D015").VALUE, String$(Len(D_rec.D015), "0"))
        D_rec.D016 = Format$(oRecordSet.Fields("D016").VALUE, String$(Len(D_rec.D016), "0"))
        D_rec.D017 = Format$(oRecordSet.Fields("D017").VALUE, String$(Len(D_rec.D017), "0"))
        D_rec.D018 = Format$(DCount, String$(Len(D_rec.D018), "0"))  '/일련번호
        D_rec.D019 = oRecordSet.Fields("D019").VALUE
        
        Print #1, MDC_SetMod.sStr(D_rec.D001) & MDC_SetMod.sStr(D_rec.D002) & MDC_SetMod.sStr(D_rec.D003) & MDC_SetMod.sStr(D_rec.D004) & MDC_SetMod.sStr(D_rec.D005) & MDC_SetMod.sStr(D_rec.D006) & MDC_SetMod.sStr(D_rec.D007) & MDC_SetMod.sStr(D_rec.D008) & MDC_SetMod.sStr(D_rec.D009) & MDC_SetMod.sStr(D_rec.D010) & _
                  MDC_SetMod.sStr(D_rec.D011) & MDC_SetMod.sStr(D_rec.D012) & MDC_SetMod.sStr(D_rec.D013) & MDC_SetMod.sStr(D_rec.D014) & MDC_SetMod.sStr(D_rec.D015) & MDC_SetMod.sStr(D_rec.D016) & MDC_SetMod.sStr(D_rec.D017) & MDC_SetMod.sStr(D_rec.D018) & MDC_SetMod.sStr(D_rec.D019)
     
        oRecordSet.MoveNext
    Loop
    
    If CheckD = False Then
        File_Create_D_record = True
    Else
        File_Create_D_record = False
    End If
    Set oRecordSet = Nothing
    Exit Function
'///////////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set oRecordSet = Nothing
   
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "D레코드가 존재하지 않습니다.", bmt_Short, smt_Error
    Else
        Call Matrix_AddRow("D레코드오류: " & Err.Description, False)
    End If
    File_Create_D_record = False
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
    ElseIf oForm.Items("TeamName").Specific.VALUE = "" Then
        ErrNum = 2
        GoTo HeaderSpaceLineDel
    ElseIf oForm.Items("Dname").Specific.VALUE = "" Then
        ErrNum = 3
        GoTo HeaderSpaceLineDel
    ElseIf oForm.Items("Dtel").Specific.VALUE = "" Then
        ErrNum = 4
        GoTo HeaderSpaceLineDel
    ElseIf oForm.Items("DocDate").Specific.VALUE = "" Then
        ErrNum = 5
        GoTo HeaderSpaceLineDel
    End If
    
    HeaderSpaceLineDel = True
    Exit Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
HeaderSpaceLineDel:
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "홈텍스ID(5자리이상)를 확인하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 2 Then
        Sbo_Application.StatusBar.SetText "담당자부서는 필수입니다. 선택하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 3 Then
        Sbo_Application.StatusBar.SetText "담당자성명은 필수입니다. 선택하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 4 Then
        Sbo_Application.StatusBar.SetText "담당자전화번호는 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 5 Then
        Sbo_Application.StatusBar.SetText "제출일자는 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
    Else
        Sbo_Application.StatusBar.SetText "HeaderSpaceLineDel 실행 중 오류가 발생했습니다." & Space$(10) & Err.Description, bmt_Short, smt_Error
    End If
    
    HeaderSpaceLineDel = False
End Function




