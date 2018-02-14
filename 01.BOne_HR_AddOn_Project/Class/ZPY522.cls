VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ZPY522"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//  SAP MANAGE UI API 2004 SDK Sample
'//****************************************************************************
'//  File           : ZPY522.cls
'//  Module         : 인사관리>정산관리
'//  Desc           : 의료비 기부금 전산매체수록
'//  FormType       : 2000060522
'//  Create Date    : 2006.01.27
'//  Modified Date  :
'//  Creator        : Ham Mi Kyoung
'//  Modifier       :
'//  Copyright  (c) Morning Data
'//****************************************************************************
Option Explicit

Public oFormUniqueID    As String
Public oForm            As SAPbouiCOM.Form
Private sRecordset      As SAPbobsCOM.Recordset

Private oMat1           As SAPbouiCOM.Matrix
Private Last_Item       As String                     '클래스에서 선택한 마지막 아이템 Uid값

Private oJsnYear        As String
Private DPTSTR          As String
Private DPTEND          As String
Private MSTCOD          As String
Private oFilePath       As String

Private FILNAM      As String * 30    '파  일  명
Private MaxRow      As Long
Private NEWCNT      As Integer
Private C_MSTCOD    As String
Private C_CLTCOD    As String

Private CLTCOD      As String
Private BUSCNT      As Integer  '/ B레코드일련번호
Private BUSTOT      As Integer  '/ B레코드총갯수

'/ 기부금지급명세서 /
Private Type H_A_Record
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
        HANCOD As String * 3     '한글코드종류
        FILLER As String * 2     '공      란
        
End Type
Dim HArec As H_A_Record

Private Type H_B_Record
        RECGBN As String * 1    '레코드구분
        DTAGBN As String * 2    '자료 구분
        TAXCOD As String * 3    '세무서
        BUSCNT As String * 6    '일련번호
        
        BUSNBR As String * 10   '사업자번호
        SANGHO As String * 40   '상      호
        
        CRECNT As String * 7    'C레코드수
        DRECNT As String * 7    'D레코드수
        GBUAMT As String * 13   '기부금액 총계
        GBUAM1 As String * 13   '공제대상 금액 총계
        
        RNGCOD As String * 1    '제출대상기간
        FILLER As String * 77   '공      란
End Type
Dim HBrec As H_B_Record

Private Type H_C_Record
        RECGBN As String * 1    '레코드구분
        DTAGBN As String * 2    '자료  구분
        TAXCOD As String * 3    '세  무  서
        SQNNBR As String * 6    '일련  번호
        
        BUSNBR As String * 10   '사업자번호
        PERNBR As String * 13   '주민  번호
        INTGBN As String * 1    '내외국인구분
        MSTNAM As String * 30   '성      명
        
        GBUCOD As String * 2    '유형코드
        GBUYER As String * 4    '기부년도
        GBUAMT As String * 13   '기부금액
        BEFAMT As String * 13   '전년까지 공제된 금액
        TARAMT As String * 13   '공제대상 금액
        CURAMT As String * 13   '해당년도 공제 금액
        DESAMT As String * 13   '해당년도 소멸 금액
        NEXAMT As String * 13   '다음년도 이월 금액
        GBUSEQ As String * 5    '기부금 조정명세 일련번호
        FILLER As String * 25   '공란
End Type
Dim HCrec As H_C_Record

Private Type H_D_Record
        RECGBN As String * 1    '레코드구분
        DTAGBN As String * 2    '자료  구분
        TAXCOD As String * 3    '세  무  서
        SQNNBR As String * 6    '일련  번호
        
        BUSNBR As String * 10   '사업자번호
        PERNBR As String * 13   '주민  번호
        
        GBUCOD As String * 2    '유형코드
        
        GBUNBR As String * 13   '기부처 사업자번호
        GBUNAM As String * 30   '기부처 상호
        
        GWANGE As String * 1    '관계코드
        INTGBN As String * 1    '내.외국인 구분
        FAMNAM As String * 20   '기부자 성명
        FAMPER As String * 13   '기부자 주민번호
        GBUCNT As String * 5    '연간 기부건수
        GBUAMT As String * 13   '연간 기부금액
        GBUSEQ As String * 5    '해당년도 기부명세 일련번호
        FILLER As String * 42   '공란
End Type
Dim HDrec As H_D_Record

'/ 의료비지급명세서 레코드 /
Private Type CA_Record
        RECGBN As String * 1     '레코드구분
        DTAGBN As String * 2     '자료  구분
        TAXCOD As String * 3     '세  무  서
        SQNNBR As String * 6     '일련  번호
        PRTDAT As String * 8     '제출  일자
        PRTBUS As String * 10    '자료제출자사업자번호
        HOMTID As String * 20    '홈텍스ID
        PGMCOD As String * 4     '세무프로그램코드
        BUSNBR As String * 10    '사업자번호
        SANGHO As String * 40    '상      호
        PERNBR As String * 13    '주민  번호
        INTGBN As String * 1     '내외국인구분
        MSTNAM As String * 30    '성      명
        JIGBUS As String * 10    '지급처사업자번호
        JIGNAM As String * 40    '지급처상호
        MEDCOD As String * 1     '의료비증빙코드(2008.12추가)
        JIGCNT As String * 5     '지급(현금)건수
        JIGAMT As String * 11    '지급(현금)금액
        JIGPER As String * 13    '지급 주민번호
        JIGINT As String * 1     '지급 내외국인구분
        JIGCHK As String * 1     '지급 해당여부
        RNGCOD As String * 1     '제출대상기간(2009년추가)
        FILD01 As String * 19    '공란(2009년추가)
End Type
Dim CArec As CA_Record

'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error
    Dim i           As Long
    Dim oXmlDoc             As New MSXML2.DOMDocument

    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\ZPY522.srf")
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
    
    oFormUniqueID = "ZPY522_" & GetTotalFormsCount
    
    '폼 할당
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)
    
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '컬렉션에 폼을 담는다   **컬렉션이란 개체를 담아 놓는 배열로서 여기서는 활성화되어져 있는 폼을 담고 있다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    AddForms Me, oFormUniqueID, "ZPY522"
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
                If pval.ItemUid = "CBtn1" Then   '/ ChooseBtn사원리스트
                    oForm.Items("MSTCOD").CLICK ct_Regular
                    Sbo_Application.ActivateMenuItem ("7425")
                    BubbleEvent = False
                ElseIf pval.ItemUid = "1" And oForm.Mode = fm_ADD_MODE Then
                    If HeaderSpaceLineDel = False Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                    If oForm.Items("JsnGbn").Specific.Selected.Value = "1" Then
                        If File_MED_Create = False Then
                            BubbleEvent = False
                            Exit Sub
                        Else
                            BubbleEvent = False
                            oForm.Mode = fm_OK_MODE
                        End If
                    Else
                        If File_GBU_Create = False Then
                            BubbleEvent = False
                            Exit Sub
                        Else
                            BubbleEvent = False
                            oForm.Mode = fm_OK_MODE
                        End If
                    
                    End If
                ElseIf pval.ItemUid = "Btn1" Then
                   oFilePath = ZP_Form.vbGetBrowseDirectory(ZP_Form)
                   oForm.DataSources.UserDataSources("Path").ValueEx = oFilePath
                   BubbleEvent = False
                   Exit Sub

               End If
            Else
            End If
'et_VALIDATE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_VALIDATE
            If pval.BeforeAction = False And pval.ItemChanged = True And _
              (pval.ItemUid = "JsnYear" Or pval.ItemUid = "MSTCOD") Then
               FlushToItemValue pval.ItemUid
            End If
'et_CLICK''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_CLICK
             If pval.BeforeAction = True And pval.ItemUid <> "1000001" And pval.ItemUid <> "2" Then
                If Last_Item = "JsnYear" Then     '/정산년도
                    If Trim$(oForm.Items(Last_Item).Specific.Value) <> "" Then
                        If MDC_SetMod.ChkYearMonth(Trim$(CStr((oForm.Items(Last_Item).Specific.Value))) & "01") = False Then
                            oForm.Items(Last_Item).Update
                            Sbo_Application.StatusBar.SetText "정산년도를 확인하여 주십시오.", bmt_Short, smt_Error
                            BubbleEvent = False
                        End If
                    End If
                ElseIf Last_Item = "MSTCOD" Then
                    If Trim$(oForm.Items(Last_Item).Specific.String) <> "" And _
                    MDC_SetMod.Value_ChkYn("[@PH_PY001A]", "Code", "'" & Trim$(oForm.Items(Last_Item).Specific.String) & "'", "") = True Then
                        oForm.Items(Last_Item).Update
                        Sbo_Application.StatusBar.SetText "사원번호를 확인하여 주십시오.", bmt_Short, smt_Error
                        BubbleEvent = False
                    End If
                End If
            End If
'et_KEY_DOWN''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_KEY_DOWN
             If pval.BeforeAction = True And pval.ItemUid = "JsnYear" And pval.CharPressed = 9 Then
                If Len(Trim$(oForm.Items(pval.ItemUid).Specific.String)) < 4 Then
                    oForm.Items(pval.ItemUid).Specific.Value = Format$(oForm.Items(pval.ItemUid).Specific.Value, "2000")
                End If
                If MDC_SetMod.ChkYearMonth(Trim$(CStr((oForm.Items(pval.ItemUid).Specific.Value))) & "01") = False Then
                    Sbo_Application.StatusBar.SetText "정산년도를 확인하여 주십시오.", bmt_Short, smt_Error
                    BubbleEvent = False
                End If
             ElseIf pval.BeforeAction = True And pval.ItemUid = "MSTCOD" And pval.CharPressed = 9 Then
                If Trim$(oForm.Items("MSTCOD").Specific.String) <> "" And _
                    MDC_SetMod.Value_ChkYn("[@PH_PY001A]", "Code", "'" & Trim$(oForm.Items("MSTCOD").Specific.String) & "'", "") = True Then
                    Sbo_Application.StatusBar.SetText "사원번호를 확인하여 주십시오.", bmt_Short, smt_Error
                    BubbleEvent = False
                End If
            End If
'et_GOT_FOCUS''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_GOT_FOCUS
            If Last_Item = "Mat1" Then
                If pval.Row > 0 Then
                    Last_Item = pval.ItemUid
                End If
            Else
                Last_Item = pval.ItemUid
            End If
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
    Dim oCombo1         As SAPbouiCOM.ComboBox
    Dim oCombo2         As SAPbouiCOM.ComboBox
    Dim oOption         As SAPbouiCOM.OptionBtn
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim oEdit           As SAPbouiCOM.EditText
    Dim oColumn         As SAPbouiCOM.Column
    Dim sQry            As String
    Dim Mat1            As SAPbouiCOM.Matrix

    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
        
    Call oForm.DataSources.UserDataSources.Add("JsnYear", dt_SHORT_TEXT, 4)    '/ 생성년도
    Call oForm.DataSources.UserDataSources.Add("JsnGbn", dt_SHORT_TEXT, 10)     '/ 생성구분
    Call oForm.DataSources.UserDataSources.Add("JSNTYP", dt_SHORT_TEXT, 10)     '/ 기간코드
    Call oForm.DataSources.UserDataSources.Add("CLTCOD", dt_SHORT_TEXT, 10)      '/ 지점
    Call oForm.DataSources.UserDataSources.Add("DptStr", dt_SHORT_TEXT, 10)     '/ 부서코드
    Call oForm.DataSources.UserDataSources.Add("DptEnd", dt_SHORT_TEXT, 10)
    Call oForm.DataSources.UserDataSources.Add("MSTCOD", dt_SHORT_TEXT, 8)
    Call oForm.DataSources.UserDataSources.Add("MSTNAM", dt_SHORT_TEXT, 30)
    Call oForm.DataSources.UserDataSources.Add("EmpID", dt_SHORT_TEXT, 10)
    Call oForm.DataSources.UserDataSources.Add("PRTDAT", dt_DATE, 10)
    Call oForm.DataSources.UserDataSources.Add("Path", dt_SHORT_TEXT)
    
    Set oEdit = oForm.Items("JsnYear").Specific
    oEdit.DataBind.SetBound True, "", "JsnYear"
    Set oEdit = oForm.Items("MSTCOD").Specific
    oEdit.DataBind.SetBound True, "", "MSTCOD"
    Set oEdit = oForm.Items("MSTNAM").Specific
    oEdit.DataBind.SetBound True, "", "MSTNAM"
    Set oEdit = oForm.Items("EmpID").Specific
    oEdit.DataBind.SetBound True, "", "EmpID"
    Set oEdit = oForm.Items("Path").Specific
    oEdit.DataBind.SetBound True, "", "Path"
    Set oEdit = oForm.Items("PRTDAT").Specific
    oEdit.DataBind.SetBound True, "", "PRTDAT"
    
    '// 생성구분
    Set oCombo1 = oForm.Items("JsnGbn").Specific
    oCombo1.ValidValues.Add "1", "의료비 지급명세서"
    oCombo1.ValidValues.Add "2", "기부금 지급명세서"
    Call oCombo1.Select(0, psk_Index)   '/ 전체
    '// 생성구분
    Set oCombo1 = oForm.Items("JSNTYP").Specific
    oCombo1.ValidValues.Add "1", "연간(01.01~12.31)지급분"
    oCombo1.ValidValues.Add "2", "폐업에 의한 수시 제출분"
    oCombo1.ValidValues.Add "3", "수시 분할제출분"
    Call oCombo1.Select(0, psk_Index)   '/ 전체
    
    '// 자료 제출자
    Set oCombo1 = oForm.Items("CLTCOD").Specific
'    sQry = " SELECT T0.U_WCHCLT, MAX(T1.NAME)  FROM [@PH_PY005A] T0 INNER JOIN [@PH_PY005A] T1 ON T0.U_WCHCLT = T1.CODE GROUP BY T0.U_WCHCLT  ORDER BY T0.U_WCHCLT"
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
    oRecordSet.DoQuery sQry
    Do Until oRecordSet.EOF
        oCombo1.ValidValues.Add Trim$(oRecordSet.Fields(0).Value), Trim$(oRecordSet.Fields(1).Value)
        oRecordSet.MoveNext
    Loop
    If oCombo1.ValidValues.Count > 0 Then
        Call oCombo1.Select(0, psk_Index)  '/ 전체
    End If
    '// 부서
    Set oCombo1 = oForm.Items("DptStr").Specific
    Set oCombo2 = oForm.Items("DptEnd").Specific
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '1' AND U_UseYN= 'Y'"
    oRecordSet.DoQuery sQry
    oCombo1.ValidValues.Add "-1", "모두"
    oCombo2.ValidValues.Add "-1", "모두"
    Do Until oRecordSet.EOF
        oCombo1.ValidValues.Add Trim$(oRecordSet.Fields(0).Value), Trim$(oRecordSet.Fields(1).Value)
        oCombo2.ValidValues.Add Trim$(oRecordSet.Fields(0).Value), Trim$(oRecordSet.Fields(1).Value)
        oRecordSet.MoveNext
    Loop
    Call oCombo1.Select(0, psk_Index)
    Call oCombo2.Select(0, psk_Index)

    Set oMat1 = oForm.Items("Mat1").Specific
    
    Call oForm.DataSources.UserDataSources.Add("Col0", dt_SHORT_TEXT)
    Call oForm.DataSources.UserDataSources.Add("Col1", dt_SHORT_TEXT)
    
    Set oColumn = oMat1.Columns("Col0")
    oColumn.DataBind.SetBound True, "", "Col0"
    
    Set oColumn = oMat1.Columns("Col1")
    oColumn.DataBind.SetBound True, "", "Col1"

    Set oEdit = Nothing
    Set oOption = Nothing
    Set oCombo1 = Nothing
    Set oCombo2 = Nothing
    Set oRecordSet = Nothing
    Set Mat1 = Nothing
    Exit Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Error_Message:
    Set oEdit = Nothing
    Set oOption = Nothing
    Set oCombo1 = Nothing
    Set oCombo2 = Nothing
    Set oRecordSet = Nothing
    Set Mat1 = Nothing
    Sbo_Application.StatusBar.SetText "CreateItems 실행 중 오류가 발생했습니다." & Space$(10) & Err.Description, bmt_Short, smt_Error
End Sub
Private Function File_GBU_Create() As Boolean
On Error GoTo Error_Message
    Dim ErrNum          As Integer
    Dim oStr            As String
    Dim sQry            As String
    
    Set sRecordset = oCompany.GetBusinessObject(BoRecordset)
    
    ErrNum = 0
    '/ Question
    If Sbo_Application.MessageBox("전산매체신고 파일을 생성하시겠습니까?", 2, "&Yes!", "&No") = 2 Then
        ErrNum = 1
        GoTo Error_Message
    End If
    
    oMat1.Clear
    MaxRow = 0
    
    BUSCNT = 0   '/ B레코드 일련번호
    BUSTOT = 0
    oJsnYear = oForm.Items("JsnYear").Specific.Value
    '/ 파일경로설정
    If oFilePath = "" Then oFilePath = "C:\EOSDATA"
    oFilePath = IIf(Right$(oFilePath, 1) = "\", oFilePath, oFilePath & "\")
    oStr = CreateFolder(Trim$(oFilePath))
    If Trim$(oStr) <> "" Then
        ErrNum = 5
        GoTo Error_Message
    End If
    
    '/ 기부금 제출자(대리인) 레코드
    If File_GBU_Create_ARecord = False Then
        ErrNum = 2
        GoTo Error_Message
    End If
    
    Close #1
    Open FILNAM For Output As #1
        '/ A레코드: 기부금 원천징수의무자별 집계 레코드
        Print #1, MDC_SetMod.sStr(HArec.RECGBN) & MDC_SetMod.sStr(HArec.DTAGBN) & MDC_SetMod.sStr(HArec.TAXCOD) & MDC_SetMod.sStr(HArec.PRTDAT) & _
                  MDC_SetMod.sStr(HArec.RPTGBN) & MDC_SetMod.sStr(HArec.TAXAGE) & MDC_SetMod.sStr(HArec.HOMTID) & MDC_SetMod.sStr(HArec.PGMCOD) & _
                  MDC_SetMod.sStr(HArec.BUSNBR) & MDC_SetMod.sStr(HArec.SANGHO) & MDC_SetMod.sStr(HArec.DAMDPT) & MDC_SetMod.sStr(HArec.DAMNAM) & _
                  MDC_SetMod.sStr(HArec.DAMTEL) & MDC_SetMod.sStr(HArec.BUSCNT) & MDC_SetMod.sStr(HArec.HANCOD) & _
                  MDC_SetMod.sStr(HArec.FILLER)
                  
        Call Matrix_AddRow("제출자 레코드 생성 완료!", True)
        
        '/ B레코드: 기부금 집계 레코드 /***********************************************/
        sQry = "SELECT Code, U_TAXCODE, U_BUSNUM, U_CLTNAME, U_COMPRT, U_PERNUM FROM [@PH_PY005A] "
        sQry = sQry & " WHERE  U_WCHCLT = '" & oForm.Items("CLTCOD").Specific.Selected.Value & "' ORDER BY Code"
        sRecordset.DoQuery sQry
        Do Until sRecordset.EOF
            '/ B레코드: 기부금 집계 레코드 /***********************************************/
            HBrec.TAXCOD = sRecordset.Fields("U_TAXCODE").Value
            HBrec.BUSNBR = Replace(sRecordset.Fields("U_BUSNUM").Value, "-", "")
            HBrec.SANGHO = Trim$(sRecordset.Fields("U_CLTNAME").Value)
            
            CLTCOD = sRecordset.Fields(0).Value

            Select Case File_GBU_Create_Brecord
                Case 0
                    Call Matrix_AddRow(CLTCOD & "- 징수의무자의 집계 레코드 생성 완료!", True)
                    '/ C레코드: 기부금 주(현)근무처 레코드 /***********************************************/
                    NEWCNT = 1
                    If File_GBU_Create_CRecord = False Then
                        ErrNum = 4
                        GoTo Error_Message
                    End If
                    Call Matrix_AddRow(CLTCOD & "- 징수의무자의 데이터 레코드" & NEWCNT & "건 생성 완료!", True)
                Case 1
                    ErrNum = 3
                    GoTo Error_Message
                Case 2  '// 해당자사에 기부금내역자료가 없으면 B,C,D레코드 생성을 건너뜀
            
            End Select
            '/
            sRecordset.MoveNext
        Loop
    Close #1
    oForm.DataSources.UserDataSources("Path").Value = FILNAM
    Sbo_Application.StatusBar.SetText "전산매체수록이 정상적으로 완료되었습니다.", bmt_Short, smt_Success
    File_GBU_Create = True
    Set sRecordset = Nothing
    Exit Function
'///////////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set sRecordset = Nothing
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "취소하였습니다.", bmt_Short, smt_Success
    ElseIf ErrNum = 2 Then
        Sbo_Application.StatusBar.SetText "A레코드(기부금 제출자 레코드) 생성 실패.", bmt_Short, smt_Error
    ElseIf ErrNum = 3 Then
        Sbo_Application.StatusBar.SetText "B레코드(기부금 원천징수의무자별 집계 레코드) 생성 실패.", bmt_Short, smt_Error
    ElseIf ErrNum = 4 Then
        Sbo_Application.StatusBar.SetText "C레코드(기부금 주(현)근무처 레코드) 생성 실패.", bmt_Short, smt_Error
    ElseIf ErrNum = 5 Then
        Sbo_Application.StatusBar.SetText "CreateFolder Error : " & oStr, bmt_Short, smt_Error
    Else
        Sbo_Application.StatusBar.SetText "File_GBU_Create 실행 중 오류가 발생했습니다." & Space$(10) & Err.Description, bmt_Short, smt_Error
    End If
    File_GBU_Create = False
End Function
Private Function File_GBU_Create_ARecord() As Boolean
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
    '// 자사정보만 있고, 기부금자료가 없어서 C,D레코드 없이 B레코드만 생성되는 경우가 있으므로
    '// 기부금 자료가 없으면 B레코드가 생성되지 않도록 함
    sQry = "SELECT COUNT(CODE) from [@PH_PY005A] T0 " & _
           "WHERE U_WCHCLT = '" & oForm.Items("CLTCOD").Specific.Selected.Value & "' " & _
           "AND Code IN (SELECT U_CLTCOD FROM [@ZPY505H] WHERE U_JSNYER = '" & oJsnYear & "')"
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
        FILNAM = oFilePath & "H" & Mid$(BUSNUM, 1, 7) & "." & Mid$(BUSNUM, 8, 3)
       '/ A Record /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
        HArec.RECGBN = "A"                          '
        HArec.DTAGBN = "27"
        HArec.TAXCOD = oRecordSet.Fields("U_TAXCODE").Value
        HArec.PRTDAT = PRTDAT
        HArec.RPTGBN = oRecordSet.Fields("U_TaxDGbn").Value
        HArec.TAXAGE = oRecordSet.Fields("U_TaxDCode").Value
        HArec.HOMTID = Trim$(oRecordSet.Fields("U_HOMETID").Value)
        HArec.PGMCOD = "9000"
        HArec.BUSNBR = Replace(oRecordSet.Fields("U_TAXDBUS").Value, "-", "")
        HArec.SANGHO = Trim$(oRecordSet.Fields("U_TAXDNAM").Value)
        HArec.DAMDPT = Trim$(oRecordSet.Fields("U_CHGDPT").Value)
        HArec.DAMNAM = Trim$(oRecordSet.Fields("U_CHGNAME").Value)
        HArec.DAMTEL = Trim$(oRecordSet.Fields("U_CHGTEL").Value)
        HArec.BUSCNT = Format$(BUSTOT, String$(Len(HArec.BUSCNT), "0"))  '/ 원천징수의무자수
        HArec.HANCOD = "101"
        HArec.FILLER = Space$(Len(HArec.FILLER))
     
     '/ 필수입력 체크
        If Trim$(HArec.TAXCOD) = "" Then
            Call Matrix_AddRow("A레코드:세무서코드가 누락되었습니다. 입력하여 주십시오.", True, True)
            CheckA = True
        End If
        If Trim$(HArec.RPTGBN) = "" Then
            Call Matrix_AddRow("A레코드:제출자구분가 누락되었습니다. 입력하여 주십시오.", True, True)
            CheckA = True
        End If
        If Trim$(HArec.BUSNBR) = "" Then
            Call Matrix_AddRow("A레코드:제출자사업자번호가 누락되었습니다. 입력하여 주십시오.", True, True)
            CheckA = True
        End If
        If Trim$(HArec.SANGHO) = "" Then
            Call Matrix_AddRow("A레코드:제출자상호가 누락되었습니다. 입력하여 주십시오.", True, True)
            CheckA = True
        End If
        If Trim$(HArec.DAMDPT) = "" Then
            Call Matrix_AddRow("A레코드:담당자부서가 누락되었습니다. 입력하여 주십시오.", True, True)
            CheckA = True
        End If
        If Trim$(HArec.DAMNAM) = "" Then
            Call Matrix_AddRow("A레코드:담당자성명이 누락되었습니다. 입력하여 주십시오.", True, True)
            CheckA = True
        End If
        If Trim$(HArec.DAMTEL) = "" Then
            Call Matrix_AddRow("A레코드:담당자전화번호가 누락되었습니다. 입력하여 주십시오.", True, True)
            CheckA = True
        End If
        If HArec.BUSCNT = 0 Then
            Call Matrix_AddRow("A레코드:신고내역이 존재하는 B레코드가 없습니다. 확인하여 주십시오.", True, True)
            CheckA = True
        End If
    End If
    
    If CheckA = False Then
        File_GBU_Create_ARecord = True
    Else
        File_GBU_Create_ARecord = False
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
    File_GBU_Create_ARecord = False
End Function

'//------------------------------------------------------------
'// 반환값 리스트
'// 0 : 정상적으로 레코드 생성
'// 1 : 에러
'// 2 : B ~ E 레코드 생성 안함
'//------------------------------------------------------------
Private Function File_GBU_Create_Brecord() As Integer
On Error GoTo Error_Message
    Dim ErrNum          As Integer
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim CheckB          As String
    
    CheckB = False  '/체크필요유무
    ErrNum = 0
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '/ 집계정보
    sQry = "EXEC ZPY522 'B', '" & oJsnYear & "', '" & CLTCOD & "', '" & _
                                  DPTSTR & "', '" & DPTEND & "','" & MSTCOD & "'"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
        ErrNum = 1
        GoTo Error_Message
    ElseIf oRecordSet.Fields("CRECNT").Value = 0 Then
        ErrNum = 2
        GoTo Error_Message
    Else
     '/ B Record /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
        BUSCNT = BUSCNT + 1
        
        HBrec.RECGBN = "B"
        HBrec.DTAGBN = "27"
        HBrec.BUSCNT = Format$(BUSCNT, String$(Len(HBrec.BUSCNT), "0"))   '/ 원천징수의무자수 일련번호
        
        HBrec.CRECNT = Format$(oRecordSet.Fields("CRECNT").Value, String$(Len(HBrec.CRECNT), "0"))    'C Record수
        HBrec.DRECNT = Format$(oRecordSet.Fields("DRECNT").Value, String$(Len(HBrec.DRECNT), "0"))    'D Record수
        HBrec.GBUAMT = Format$(oRecordSet.Fields("GBUAMT").Value, String$(Len(HBrec.GBUAMT), "0"))    '기부금액 총계
        HBrec.GBUAM1 = Format$(oRecordSet.Fields("GBUAM1").Value, String$(Len(HBrec.GBUAM1), "0"))    '공제대상 금액 총계
        HBrec.RNGCOD = oForm.Items("JSNTYP").Specific.Selected.Value
        HBrec.FILLER = Space$(Len(HBrec.FILLER))
        
        Print #1, MDC_SetMod.sStr(HBrec.RECGBN) & MDC_SetMod.sStr(HBrec.DTAGBN) & MDC_SetMod.sStr(HBrec.TAXCOD) & _
            MDC_SetMod.sStr(HBrec.BUSCNT) & MDC_SetMod.sStr(HBrec.BUSNBR) & MDC_SetMod.sStr(HBrec.SANGHO) & _
            MDC_SetMod.sStr(HBrec.CRECNT) & MDC_SetMod.sStr(HBrec.DRECNT) & MDC_SetMod.sStr(HBrec.GBUAMT) & _
            MDC_SetMod.sStr(HBrec.GBUAM1) & MDC_SetMod.sStr(HBrec.RNGCOD) & MDC_SetMod.sStr(HBrec.FILLER)
                  
     '/ 필수입력 체크
        If Trim$(HBrec.BUSNBR) = "" Then
            Call Matrix_AddRow("B레코드:사업자번호가 누락되었습니다. 입력하여 주십시오.", True, True)
            CheckB = True
        End If
    End If
    
    If CheckB = False Then
        File_GBU_Create_Brecord = 0
    Else
        File_GBU_Create_Brecord = 1
    End If
    Set oRecordSet = Nothing
    Exit Function
'///////////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set oRecordSet = Nothing
   
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "집계레코드가 존재하지 않습니다. 등록하여 주십시오.", bmt_Short, smt_Error
        File_GBU_Create_Brecord = 1
    ElseIf ErrNum = 2 Then
        File_GBU_Create_Brecord = 2
    Else
        Call Matrix_AddRow("B레코드오류: " & Err.Description, False)
        File_GBU_Create_Brecord = 1
    End If
    
End Function
Private Function File_GBU_Create_CRecord() As Boolean
On Error GoTo Error_Message
    Dim ErrNum      As Integer
    Dim oRecordSet  As SAPbobsCOM.Recordset
    Dim sQry        As String
    Dim CheckC      As String
    Dim BefMSTCOD   As String:  Dim RecCNT      As Long
    
    CheckC = False  '/체크필요유무
    
    ErrNum = 0
    RecCNT = 0
    C_MSTCOD = ""
    C_CLTCOD = ""
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

'/ 사원 정보 조회
    
    BefMSTCOD = ""
    sQry = "EXEC ZPY522 'C', '" & oJsnYear & "', '" & CLTCOD & "', '" & _
                                  DPTSTR & "', '" & DPTEND & "','" & MSTCOD & "'"
 
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
'        ErrNum = 1
'        GoTo error_Message
    End If
    
    Do Until oRecordSet.EOF
    
        '/ D레코드: 해당년도 기부명세 레코드
        If BefMSTCOD <> "" And BefMSTCOD <> Trim$(oRecordSet.Fields("MSTCOD").Value) Then
            If File_GBU_Create_DRecord = False Then
                ErrNum = 2
                GoTo Error_Message
            End If
            NEWCNT = NEWCNT + 1
            RecCNT = 0
        End If
    
        C_MSTCOD = Trim$(oRecordSet.Fields("MSTCOD").Value)
        C_CLTCOD = Trim$(oRecordSet.Fields("CLTCOD").Value)

    '/ 기부금 기부금 조정명세 레코드 /
        '/ C Record /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        HCrec.RECGBN = "C"                                                                          '/ 레코드구분
        HCrec.DTAGBN = "27"                                                                         '/ 자료구분
        HCrec.TAXCOD = HBrec.TAXCOD                                                                 '/ 세무서
        HCrec.SQNNBR = Format$(NEWCNT, String$(Len(HCrec.SQNNBR), "0"))                               '/ 일련번호
        
        HCrec.BUSNBR = HBrec.BUSNBR                                                                 '/ 원천징수의무자 사업자번호
        HCrec.PERNBR = Replace(oRecordSet.Fields("PERNBR").Value, "-", "")                          '/ 소득자 주민번호
        HCrec.INTGBN = Format$(oRecordSet.Fields("INTGBN").Value, String$(Len(HCrec.INTGBN), "0"))    '/ 내.외국인 구분
        HCrec.MSTNAM = oRecordSet.Fields("MSTNAM").Value                                            '/ 성명
        
        HCrec.GBUCOD = oRecordSet.Fields("GBUCOD").Value                                            '/ 유형코드
        HCrec.GBUYER = Format$(oRecordSet.Fields("GBUYER").Value, String$(Len(HCrec.GBUYER), "0"))    '/ 기부년도
        HCrec.GBUAMT = Format$(oRecordSet.Fields("GBUAMT").Value, String$(Len(HCrec.GBUAMT), "0"))   '/ 기부금액
        HCrec.BEFAMT = Format$(oRecordSet.Fields("BEFAMT").Value, String$(Len(HCrec.BEFAMT), "0"))   '/ 전년까지 공제된 금액
        HCrec.TARAMT = Format$(oRecordSet.Fields("TARAMT").Value, String$(Len(HCrec.TARAMT), "0"))   '/ 공제대상금액
        HCrec.CURAMT = Format$(oRecordSet.Fields("CURAMT").Value, String$(Len(HCrec.CURAMT), "0"))   '/ 당해년도 공제금액
        HCrec.DESAMT = Format$(oRecordSet.Fields("DESAMT").Value, String$(Len(HCrec.DESAMT), "0"))   '/ 당해년도 소멸금액
        HCrec.NEXAMT = Format$(oRecordSet.Fields("NEXAMT").Value, String$(Len(HCrec.NEXAMT), "0"))   '/ 당해년도 이월금액
        RecCNT = RecCNT + 1
        HCrec.GBUSEQ = Format$(RecCNT, String$(Len(HCrec.GBUSEQ), "0"))                              '/ 소득자별 레코드 순번
        HCrec.FILLER = Space$(Len(HCrec.FILLER))
        
        Print #1, MDC_SetMod.sStr(HCrec.RECGBN) & MDC_SetMod.sStr(HCrec.DTAGBN) & MDC_SetMod.sStr(HCrec.TAXCOD) & _
            MDC_SetMod.sStr(HCrec.SQNNBR) & MDC_SetMod.sStr(HCrec.BUSNBR) & MDC_SetMod.sStr(HCrec.PERNBR) & _
            MDC_SetMod.sStr(HCrec.INTGBN) & MDC_SetMod.sStr(HCrec.MSTNAM) & MDC_SetMod.sStr(HCrec.GBUCOD) & _
            MDC_SetMod.sStr(HCrec.GBUYER) & MDC_SetMod.sStr(HCrec.GBUAMT) & MDC_SetMod.sStr(HCrec.BEFAMT) & _
            MDC_SetMod.sStr(HCrec.TARAMT) & MDC_SetMod.sStr(HCrec.CURAMT) & MDC_SetMod.sStr(HCrec.DESAMT) & _
            MDC_SetMod.sStr(HCrec.NEXAMT) & MDC_SetMod.sStr(HCrec.GBUSEQ) & MDC_SetMod.sStr(HCrec.FILLER)
           
              
        Call Matrix_AddRow("C레코드:" & C_MSTCOD & Trim$(HCrec.MSTNAM) & " 생성 완료.", True)
     '/ 필수입력 체크
        If HCrec.INTGBN = "1" And Len(Trim$(HCrec.PERNBR)) <> 13 Then
            Call Matrix_AddRow("C레코드:" & C_MSTCOD & Trim$(HCrec.MSTNAM) & "-주민등록번호를 확인하여 주십시오.", False, True)
            CheckC = True
        End If
        
        BefMSTCOD = Trim$(oRecordSet.Fields("MSTCOD").Value)
    
        oRecordSet.MoveNext
    Loop
    '/ D레코드: 해당년도 기부명세 레코드
    If File_GBU_Create_DRecord = False Then
        ErrNum = 2
        GoTo Error_Message
    End If
        
    If CheckC = False Then
        File_GBU_Create_CRecord = True
    Else
        File_GBU_Create_CRecord = False
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
    File_GBU_Create_CRecord = False
End Function

Private Function File_GBU_Create_DRecord() As Boolean
On Error GoTo Error_Message
    Dim ErrNum      As Integer
    Dim oRecordSet  As SAPbobsCOM.Recordset
    Dim sQry        As String
    Dim CheckD      As String
    Dim RecCNT      As Integer
    
    CheckD = False  '/체크필요유무
    ErrNum = 0
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '/ 종전근무지정보
    sQry = "EXEC ZPY522 'D', '" & oJsnYear & "', '" & C_CLTCOD & "', '" & _
                                  DPTSTR & "', '" & DPTEND & "','" & C_MSTCOD & "'"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
        File_GBU_Create_DRecord = True
        Exit Function
    End If
    RecCNT = 0
    Do Until oRecordSet.EOF
        '/ D Record /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        RecCNT = RecCNT + 1       '/ 사원별 해당년도 기부명세 일련번호
        HDrec.RECGBN = "D"                                                                        '/레코드구분
        HDrec.DTAGBN = "27"                                                                       '/자료구분
        HDrec.TAXCOD = HArec.TAXCOD                                                               '/세무서
        HDrec.SQNNBR = Format$(NEWCNT, String$(Len(HDrec.SQNNBR), "0"))                             '/소득자 일련번호
        
        HDrec.BUSNBR = HBrec.BUSNBR                                                               '/원천징수의무자 사업자등록번호
        HDrec.PERNBR = HCrec.PERNBR                                                               '/소득자 주민번호
        
        HDrec.GBUCOD = oRecordSet.Fields("GBUCOD").Value                                          '/유형코드
        HDrec.GBUNBR = Trim$(Replace(oRecordSet.Fields("GBUNBR").Value, "-", ""))                  '/기부처 사업자번호
        HDrec.GBUNAM = Trim$(oRecordSet.Fields("GBUNAM").Value)                                    '/기부처명
        HDrec.GWANGE = Trim$(oRecordSet.Fields("GWANGE").Value)                                    '/기부자 관계코드
        HDrec.INTGBN = Trim$(oRecordSet.Fields("INTGBN").Value)                                    '/기부자 내.외국인
        HDrec.FAMNAM = Trim$(oRecordSet.Fields("FAMNAM").Value)                                    '/기부자명
        HDrec.FAMPER = Trim$(Replace$(oRecordSet.Fields("FAMPER").Value, "-", ""))                 '/기부자 주민번호
        HDrec.GBUCNT = Format$(oRecordSet.Fields("GBUCNT").Value, String$(Len(HDrec.GBUCNT), "0"))  '/기부 건수
        HDrec.GBUAMT = Format$(oRecordSet.Fields("GBUAMT").Value, String$(Len(HDrec.GBUAMT), "0"))  '/기부 금액
        HDrec.GBUSEQ = Format$(RecCNT, String$(Len(HDrec.GBUSEQ), "0"))                             '/소득자별 레코드 순번
        HDrec.FILLER = Space$(Len(HDrec.FILLER))
      
        Print #1, MDC_SetMod.sStr(HDrec.RECGBN) & MDC_SetMod.sStr(HDrec.DTAGBN) & MDC_SetMod.sStr(HDrec.TAXCOD) & _
                MDC_SetMod.sStr(HDrec.SQNNBR) & MDC_SetMod.sStr(HDrec.BUSNBR) & MDC_SetMod.sStr(HDrec.PERNBR) & _
                MDC_SetMod.sStr(HDrec.GBUCOD) & MDC_SetMod.sStr(HDrec.GBUNBR) & MDC_SetMod.sStr(HDrec.GBUNAM) & _
                MDC_SetMod.sStr(HDrec.GWANGE) & MDC_SetMod.sStr(HDrec.INTGBN) & MDC_SetMod.sStr(HDrec.FAMNAM) & _
                MDC_SetMod.sStr(HDrec.FAMPER) & MDC_SetMod.sStr(HDrec.GBUCNT) & MDC_SetMod.sStr(HDrec.GBUAMT) & _
                MDC_SetMod.sStr(HDrec.GBUSEQ) & MDC_SetMod.sStr(HDrec.FILLER)
        
        oRecordSet.MoveNext
    Loop
    
    Call Matrix_AddRow("D레코드:" & C_MSTCOD & Trim$(HCrec.MSTNAM) & " 당해년도 기부명세 생성 완료.", True)
    
    If CheckD = False Then
        File_GBU_Create_DRecord = True
    Else
        File_GBU_Create_DRecord = False
    End If
    Set oRecordSet = Nothing
    Exit Function
'///////////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set oRecordSet = Nothing
   
    Call Matrix_AddRow("D레코드오류: " & Err.Description, False)
    File_GBU_Create_DRecord = False
End Function

Private Function File_MED_Create() As Boolean
On Error GoTo Error_Message
    Dim ErrNum          As Integer
    Dim oStr            As String
    Dim sQry            As String
    Dim CLTNAM          As String
    
    Set sRecordset = oCompany.GetBusinessObject(BoRecordset)
    
    ErrNum = 0
    '/ Question
    If Sbo_Application.MessageBox("전산매체신고 파일을 생성하시겠습니까?", 2, "&Yes!", "&No") = 2 Then
        ErrNum = 1
        GoTo Error_Message
    End If
    
    oMat1.Clear
    MaxRow = 0
  '/
    '/ 파일경로설정
    If oFilePath = "" Then oFilePath = "C:\EOSDATA"
    oFilePath = IIf(Right$(oFilePath, 1) = "\", oFilePath, oFilePath & "\")
    oStr = CreateFolder(Trim$(oFilePath))
    If Trim$(oStr) <> "" Then
        ErrNum = 5
        GoTo Error_Message
    End If
    
    '/ 기부금 제출자(대리인) 레코드
    If File_MED_Create_ARecord = False Then
        ErrNum = 2
        GoTo Error_Message
    End If
    '/
    Close #1
    Open FILNAM For Output As #1
        sQry = "SELECT Code, U_TAXCODE, U_BUSNUM, U_CLTNAME, U_COMPRT, U_PERNUM FROM [@PH_PY005A] T0 "
        sQry = sQry & " WHERE U_WCHCLT = '" & oForm.Items("CLTCOD").Specific.Selected.Value & "' ORDER BY CODE"
        sRecordset.DoQuery sQry
        NEWCNT = 0  '// 일련번호를
        Do Until sRecordset.EOF
         '/ 원천징수의무자정보
            CArec.TAXCOD = sRecordset.Fields("U_TAXCODE").Value
            CArec.BUSNBR = Replace(sRecordset.Fields("U_BUSNUM").Value, "-", "")
            CArec.SANGHO = Trim$(sRecordset.Fields("U_CLTNAME").Value)
            CLTCOD = sRecordset.Fields(0).Value
            CLTNAM = MDC_SetMod.Get_ReData("Name", "Code", "[@PH_PY005A]", "'" & Trim$(CLTCOD) & "'")
            '/ 의료비명세 레코드
            If File_MED_Create_CARecord = False Then
                ErrNum = 3
                GoTo Error_Message
            End If
    
            Call Matrix_AddRow("신고의무자 : " & CLTCOD & "-" & CLTNAM & " 의 데이터 레코드" & NEWCNT & "건 생성 완료!", True)
         '/
            sRecordset.MoveNext
        Loop
    Close #1
    oForm.DataSources.UserDataSources("Path").Value = FILNAM
    Sbo_Application.StatusBar.SetText "전산매체수록이 정상적으로 완료되었습니다.", bmt_Short, smt_Success
    File_MED_Create = True
    Set sRecordset = Nothing
    Exit Function
'///////////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
Set sRecordset = Nothing
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "취소하였습니다.", bmt_Short, smt_Success
    ElseIf ErrNum = 2 Then
        Sbo_Application.StatusBar.SetText "제출자자료 조회를 실패하였습니다.", bmt_Short, smt_Error
    ElseIf ErrNum = 3 Then
        Sbo_Application.StatusBar.SetText "의료비명세 파일생성 실패.", bmt_Short, smt_Error
    ElseIf ErrNum = 5 Then
        Sbo_Application.StatusBar.SetText "CreateFolder Error : " & oStr, bmt_Short, smt_Error
    Else
        Sbo_Application.StatusBar.SetText "File_MED_Create 실행 중 오류가 발생했습니다." & Space$(10) & Err.Description, bmt_Short, smt_Error
    End If
    File_MED_Create = False
End Function

Private Function File_MED_Create_ARecord() As Boolean
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
    sQry = "SELECT COUNT(CODE) from [@PH_PY005A] T0 WHERE U_WCHCLT = '" & oForm.Items("CLTCOD").Specific.Selected.Value & "'"
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
        BUSNUM = Replace(oRecordSet.Fields("U_BUSNUM").Value, "-", "")              '/ 징수의무자서업자번호
        If Len(Trim$(BUSNUM)) <> 10 Then
            ErrNum = 2
            GoTo Error_Message
        End If
        CArec.PRTBUS = Replace(oRecordSet.Fields("U_TAXDBUS").Value, "-", "")       '제출자사업자번호
        CArec.PRTDAT = PRTDAT                                                       '제출  일자
        CArec.HOMTID = Trim$(oRecordSet.Fields("U_HOMETID").Value)                   '홈텍스ID
        CArec.PGMCOD = "9000"                                                       '세무프로그램코드
     '/ 필수입력 체크
        If Trim$(CArec.TAXCOD) = "" Then
            Call Matrix_AddRow("세무서코드가 누락되었습니다. 입력하여 주십시오.", True, True)
            CheckA = True
        End If
        If Trim$(CArec.BUSNBR) = "" Then
            Call Matrix_AddRow("징수의무자의 사업자번호가 누락되었습니다. 입력하여 주십시오.", True, True)
            CheckA = True
        End If
        If Trim$(CArec.PRTBUS) = "" Then
            Call Matrix_AddRow("자료제출자의 사업자번호가 누락되었습니다. 입력하여 주십시오.", True, True)
            CheckA = True
        End If
        If Trim$(CArec.SANGHO) = "" Then
            Call Matrix_AddRow("상호가 누락되었습니다. 입력하여 주십시오.", True, True)
            CheckA = True
        End If
        FILNAM = oFilePath & "CA" & Mid$(BUSNUM, 1, 7) & "." & Mid$(BUSNUM, 8, 3)
    End If
    
    If CheckA = False Then
        File_MED_Create_ARecord = True
    Else
        File_MED_Create_ARecord = False
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
    File_MED_Create_ARecord = False
End Function


Private Function File_MED_Create_CARecord() As Boolean
On Error GoTo Error_Message
    Dim ErrNum          As Integer
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim CheckC          As String
    CheckC = False  '/체크필요유무
    ErrNum = 0
    C_MSTCOD = ""
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '/ 사원별 의료비명세정보
    sQry = " SELECT T0.*, "
    sQry = sQry & " T1.U_MSTCOD, "
    sQry = sQry & " T1.U_MSTNAM, "
    sQry = sQry & " ISNULL(T4.U_govID,'') AS MSTPER,"
    sQry = sQry & " ISNULL(T4.U_INTGBN,1)   AS MSTINT"
    sQry = sQry & " FROM [@ZPY506L] T0 "
    sQry = sQry & "      INNER JOIN [@ZPY506H] T1 ON T0.DocEntry = T1.DocEntry"
    sQry = sQry & "      INNER JOIN [@PH_PY001A] T4 ON T1.U_MSTCOD = T4.Code"
    sQry = sQry & "      INNER JOIN [@ZPY504H] T5 ON T1.U_MSTCOD = T5.U_MSTCOD AND T1.U_JSNYER = T5.U_JSNYER AND T1.U_CLTCOD = T5.U_CLTCOD"
    sQry = sQry & " WHERE   T1.U_JSNYER = '" & oJsnYear & "'"
    sQry = sQry & " AND     ISNULL(T1.U_CLTCOD, '') = '" & Trim$(CLTCOD) & "'"
    sQry = sQry & " AND     T4.U_TeamCode BETWEEN " & "'" & DPTSTR & "'" & " And " & "'" & DPTEND & "'"
    sQry = sQry & " AND     T1.U_MSTCOD LIKE " & "N'" & Trim$(MSTCOD) & "'"
    sQry = sQry & " AND     ISNULL(T5.U_PILMED,0) >= 2000000"
    sQry = sQry & " ORDER BY T0.DocEntry"
    oRecordSet.DoQuery sQry
'    If oRecordSet.RecordCount = 0 Then  '// 신고자의 원천신고의무자가 1건이상일경우 기부금이 0건인 의무자자료 존재가능성있어서..주석처리함.
'        ErrNum = 1
'        GoTo error_Message
'    End If
    Do Until oRecordSet.EOF
        NEWCNT = NEWCNT + 1
        C_MSTCOD = oRecordSet.Fields("U_MSTCOD").Value
        '/ CA Record /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
        CArec.RECGBN = "A"
        CArec.DTAGBN = "26"
        CArec.SQNNBR = Format$(NEWCNT, String$(Len(CArec.SQNNBR), "0"))  '/ 일련번호
        
        CArec.PERNBR = Replace(oRecordSet.Fields("MSTPER").Value, "-", "")
        CArec.INTGBN = oRecordSet.Fields("MSTINT").Value
        CArec.MSTNAM = oRecordSet.Fields("U_MSTNAM").Value
        CArec.JIGBUS = Replace(oRecordSet.Fields("U_MEDNBR").Value, "-", "")
        CArec.JIGNAM = oRecordSet.Fields("U_MEDNAM").Value
        CArec.MEDCOD = oRecordSet.Fields("U_MEDCOD").Value
'        CArec.CADCNT = Format$(oRecordSet.Fields("U_MEDCNT2").Value, String$(Len(CArec.CADCNT), "0"))
'        CArec.CADAMT = Format$(oRecordSet.Fields("U_MEDAMT2").Value, String$(Len(CArec.CADAMT), "0"))
'        CArec.JIGCOD = oRecordSet.Fields("U_GWANGE").Value
        CArec.JIGCNT = Format$(oRecordSet.Fields("U_MEDCNT").Value, String$(Len(CArec.JIGCNT), "0"))
        CArec.JIGAMT = Format$(oRecordSet.Fields("U_MEDAMT").Value, String$(Len(CArec.JIGAMT), "0"))

        CArec.JIGPER = Replace(oRecordSet.Fields("U_PERNBR").Value, "-", "")
        CArec.JIGINT = oRecordSet.Fields("U_INTGBN").Value
        CArec.JIGCHK = oRecordSet.Fields("U_DAECHK").Value
        CArec.FILD01 = Space$(Len(CArec.FILD01))
        CArec.RNGCOD = oForm.Items("JSNTYP").Specific.Selected.Value
                                
        
        Print #1, CArec.RECGBN & CArec.DTAGBN & CArec.TAXCOD & CArec.SQNNBR & CArec.PRTDAT & CArec.PRTBUS & CArec.HOMTID & CArec.PGMCOD & _
                  CArec.BUSNBR & MDC_SetMod.sStr(CArec.SANGHO) & CArec.PERNBR & CArec.INTGBN & MDC_SetMod.sStr(CArec.MSTNAM) & CArec.JIGBUS & MDC_SetMod.sStr(CArec.JIGNAM) & _
                  CArec.MEDCOD & CArec.JIGCNT & CArec.JIGAMT & CArec.JIGPER & CArec.JIGINT & CArec.JIGCHK & CArec.RNGCOD & CArec.FILD01
        
        
       '/ 필수입력 체크
        If Trim$(CArec.JIGBUS) = "" And Trim$(CArec.MEDCOD) <> "1" Then
            Call Matrix_AddRow(C_MSTCOD & Trim$(CArec.MSTNAM) & "지급처사업자등록번호가 누락되었습니다.", True, True)
            CheckC = True
        End If
        If Trim$(CArec.MEDCOD) = "" Then
            Call Matrix_AddRow(C_MSTCOD & Trim$(CArec.MSTNAM) & "증빙코드가 누락되었습니다.", True, True)
            CheckC = True
        End If
        If Trim$(CArec.JIGPER) = "" Then
            Call Matrix_AddRow(C_MSTCOD & Trim$(CArec.MSTNAM) & "지급대상자주민등록번호가 누락되었습니다.", True, True)
            CheckC = True
        End If
        
        oRecordSet.MoveNext
    Loop
        
        
    If CheckC = False Then
        File_MED_Create_CARecord = True
    Else
        File_MED_Create_CARecord = False
    End If
    Set oRecordSet = Nothing
    Exit Function
'///////////////////////////////////////////////////////////////////////////////////////////////////////
Error_Message:
    Set oRecordSet = Nothing
   
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "의료비명세자료 존재하지 않습니다. 등록하여 주십시오.", bmt_Short, smt_Error
        Call Matrix_AddRow("CA레코드오류: " & CArec.BUSNBR & "사업장의 의료비명세자료 존재하지 않습니다.", False)
    Else
        Call Matrix_AddRow("CA레코드오류: " & Err.Description, False)
    End If
    File_MED_Create_CARecord = False
End Function

Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long)
    Dim MstInfo         As ZPAY_g_EmpID
        
    Select Case oUID
      Case "JsnYear"
            oJsnYear = oForm.Items(oUID).Specific.String
            If Trim$(oJsnYear) = "" Then
                ZPAY_GBL_JSNYER = oJsnYear
            End If
      Case "MSTCOD"
            If oForm.Items(oUID).Specific.String = "" Then
                oForm.Items(oUID).Specific.String = ""
                oForm.DataSources.UserDataSources("MSTNAM").ValueEx = ""
                oForm.DataSources.UserDataSources("EmpID").ValueEx = ""
            Else
                oForm.Items(oUID).Specific.String = UCase$(oForm.Items(oUID).Specific.String)
                MstInfo = MDC_SetMod.Get_EmpID_InFo(oForm.Items(oUID).Specific.String)
                oForm.DataSources.UserDataSources("MSTNAM").ValueEx = MstInfo.MSTNAM
                oForm.DataSources.UserDataSources("EmpID").ValueEx = MstInfo.EmpID
            End If
            oForm.Items("MSTNAM").Update
            oForm.Items("EmpID").Update
    End Select
    oForm.Items(oUID).Update
End Sub

Private Function HeaderSpaceLineDel() As Boolean
'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
'저장할 데이터의 유효성을 점검한다
'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
On Error GoTo HeaderSpaceLineDel
    Dim ErrNum      As Integer
    
    ErrNum = 0
    '/ 필수Check
    If MDC_SetMod.ChkYearMonth(oForm.Items("JsnYear").Specific.String & "01") = False Then     '/ 정산년도
        ErrNum = 1
        GoTo HeaderSpaceLineDel
    ElseIf oForm.Items("JsnGbn").Specific.Selected Is Nothing Then
        ErrNum = 2
        GoTo HeaderSpaceLineDel
    ElseIf oForm.Items("CLTCOD").Specific.Selected Is Nothing Then
        ErrNum = 3
        GoTo HeaderSpaceLineDel
    ElseIf Trim$(oForm.Items("PRTDAT").Specific.String) = "" Then
        ErrNum = 4
        GoTo HeaderSpaceLineDel
    End If
    
    DPTSTR = oForm.Items("DptStr").Specific.Selected.Value
    If DPTSTR = "-1" Then DPTSTR = "00000001"
    DPTEND = oForm.Items("DptEnd").Specific.Selected.Value
    If DPTEND = "-1" Then DPTEND = "ZZZZZZZZ"
    MSTCOD = oForm.Items("MSTCOD").Specific.Value
    If Trim$(MSTCOD) = "" Then MSTCOD = "%"
    
    HeaderSpaceLineDel = True
    Exit Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
HeaderSpaceLineDel:
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "귀속년도를 확인하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 2 Then
        Sbo_Application.StatusBar.SetText "기간코드는 필수입니다. 선택하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 3 Then
        Sbo_Application.StatusBar.SetText "지점은 필수입니다. 선택하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 4 Then
        Sbo_Application.StatusBar.SetText "제출일자는 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
    Else
        Sbo_Application.StatusBar.SetText "HeaderSpaceLineDel 실행 중 오류가 발생했습니다." & Space$(10) & Err.Description, bmt_Short, smt_Error
    End If
    
    HeaderSpaceLineDel = False
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

