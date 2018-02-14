Attribute VB_Name = "MDC_Com"
Option Explicit
Public Type MonthDate
    StDate As String '* 10 'YYYY-MM-DD
    EdDate As String '* 10  'YYYY-MM-DD
End Type


Public Type PerSonNo_Info
    check As Boolean
    BirthDay As String
    Sex As Integer     '남 : 0, 여 : 1
    ManAge As Integer
    Age As Integer
End Type

Public Function GetSpStr(Str As String) As String
    'Function ID : GetSpStr
    '기    능    : 한칸 뛰워진 값의 앞에값을 가져온다
    '인    수    : Str
    '반 환 값    : None
    '특이사항    : ex) Str = 0 강대봉, 리턴 = 강대봉
    Dim i As Long

    i = InStr(1, Str, " ", vbTextCompare)

    If i > 0 Then
        GetSpStr = Mid(Str, 1, i - 1)
    Else
        GetSpStr = ""
    End If
End Function

Public Function GetSpStr2(Str As String) As String
    'Function ID : GetSpStr2
    '기    능    : 한칸 뛰워진 값의 뒤에값을 가져온다
    '인    수    : Str
    '반 환 값    : None
    '특이사항    : ex) Str = 0 강대봉, 리턴 = 0
    Dim Buf As String

    Buf = GetSpStr(Str)

    If Len(Str) > Len(Buf) + 1 Then
        GetSpStr2 = Mid(Str, Len(Buf) + 2, Len(Str) - Len(Buf) - 1)
    End If
End Function

Public Function uISDATE(Dt As Variant, Conv As Variant) As Variant
    'Function ID : uISDATE
    '기    능    : 날짜형식 가부의 판정
    '인    수    : dt,Conv
    '반 환 값    : None
    '특이사항    : None
    If IsDate(Dt) Then
        uISDATE = Dt
    Else
        uISDATE = Conv
    End If
End Function

Public Function uISNULL(Str As Variant, Conv As Variant) As Variant
    'Function ID : uISNULL
    '기    능    : 문자형식 가부의 판정
    '인    수    : Str,Conv
    '반 환 값    : None
    '특이사항    : None
    uISNULL = IIf(IsNull(Str), Conv, Str)
End Function

Public Function uISNUMERIC(num As Variant, Conv As Variant, NumType As String) As Variant
    'Function ID : uISNUMERIC
    '기    능    : 숫자형식의 변환
    '인    수    : Str,Conv
    '반 환 값    : None
    '특이사항    : None
    If IsNumeric(num) = True Then
        Select Case NumType
        Case "INT"
            uISNUMERIC = CInt(num)
        Case "FLT"
            uISNUMERIC = CSng(num)
        Case "DBL"
            uISNUMERIC = CDbl(num)
        Case "CUR"
            uISNUMERIC = CCur(num)
        Case "LNG"
            uISNUMERIC = CLng(num)
        End Select
    Else
        uISNUMERIC = Conv
    End If
End Function
Function GetLeftNumZero(num As String, 자리수 As Long) As String
    'Function ID : GetLeftNumZero
    '기    능    : 숫자를 오른정렬하면서 왼쪽엔 '0'를 채우는 함수
    '인    수    : Num,자리수
    '반 환 값    : None
    '특이사항    :
    ' ▶ex) GetLeftNumZero("123456",10) --> 반환값 = "0000123456"
    Dim Ln As Long

    Ln = Len(num)
    If Ln <= 자리수 Then
        GetLeftNumZero = String(자리수 - Ln, "0") + num
    Else
        GetLeftNumZero = String(자리수, "0")
    End If
End Function

'-----------------------------------------------------------------------------------------
'   네비게이션 컨트롤 관련 보이기/감추기 함수
'   -> 미리보기, 출력, 행삭제, 찾기, 추가, 다음, 이전, 맨처음, 맨끝, 취소
'-----------------------------------------------------------------------------------------
Public Sub MDC_GP_EnableMenus(MDC_eForm As SAPbouiCOM.Form, _
                              ByVal MDC_bPreview As Boolean, _
                              ByVal MDC_bPrint As Boolean, _
                              ByVal MDC_bDeleteRow As Boolean, _
                              ByVal MDC_bFind As Boolean, _
                              ByVal MDC_bAdd As Boolean, _
                              ByVal MDC_bNextRecord As Boolean, _
                              ByVal MDC_bPreviousRecord As Boolean, _
                              ByVal MDC_bFirstRecord As Boolean, _
                              ByVal MDC_bLastRecord As Boolean, _
                              ByVal MDC_bCancel As Boolean, _
                              Optional ByVal MDC_bRowAdd As Boolean = False, _
                              Optional ByVal MDC_bDuplicate As Boolean = False, _
                              Optional ByVal MDC_bRemove As Boolean = False, _
                              Optional ByVal MDC_bRowClose As Boolean = False, _
                              Optional ByVal MDC_bClose As Boolean = False, _
                              Optional ByVal MDC_bRestore As Boolean = False)

    '//If Left(MDC_eForm.Type, 2) = "20" Then
        MDC_eForm.EnableMenu "519", MDC_bPreview         '// 인쇄[미리보기]
        MDC_eForm.EnableMenu "520", MDC_bPrint           '// 인쇄[출력]
        MDC_eForm.EnableMenu "1293", MDC_bDeleteRow      '// 행삭제
        MDC_eForm.EnableMenu "1281", MDC_bFind           '// 문서찾기
        MDC_eForm.EnableMenu "1282", MDC_bAdd            '// 문서추가
        MDC_eForm.EnableMenu "1283", MDC_bRemove         '// 문서제거(데이터 삭제시 사용)
        MDC_eForm.EnableMenu "1286", MDC_bClose          '// 문서닫기
        MDC_eForm.EnableMenu "1288", MDC_bNextRecord     '// 다음
        MDC_eForm.EnableMenu "1289", MDC_bPreviousRecord '// 이전
        MDC_eForm.EnableMenu "1290", MDC_bFirstRecord    '// 맨처음
        MDC_eForm.EnableMenu "1291", MDC_bLastRecord     '// 맨끝
        MDC_eForm.EnableMenu "1284", MDC_bCancel         '// 문서취소
        MDC_eForm.EnableMenu "1292", MDC_bRowAdd         '// 행추가
        MDC_eForm.EnableMenu "1287", MDC_bDuplicate      '// 문서복제
        MDC_eForm.EnableMenu "1299", MDC_bRowClose       '// 행닫기
        MDC_eForm.EnableMenu "1285", MDC_bRestore
    '//End If
End Sub

'// Matrix Combo Box Setting
Public Sub MDC_GP_MatrixSetMatComboList(MDC_fCombo As SAPbouiCOM.Column, _
                                        MDC_fSQL As String, _
                                        Optional AndLine$, _
                                        Optional AddSpace$)
    'Function ID : GetListIndex
    '기    능    :
    '인    수    : Lst
    '반 환 값    : None
    '특이사항    : 콤보박스의 들어가야 할 내용을 시스템 코드에서 가져와 세팅한다
    Dim MDC_fRecordset As SAPbobsCOM.Recordset
    
    Set MDC_fRecordset = oCompany.GetBusinessObject(BoRecordset)
    MDC_fRecordset.DoQuery MDC_fSQL

    If AddSpace = "Y" Then
        Call MDC_fCombo.ValidValues.Add("", "")
    End If
    Do Until MDC_fRecordset.EOF
        Call MDC_fCombo.ValidValues.Add(MDC_fRecordset.Fields(0).VALUE, MDC_fRecordset.Fields(1).VALUE)
        MDC_fRecordset.MoveNext
    Loop
        
    Set MDC_fRecordset = Nothing
    
End Sub
Public Sub MDC_GP_MatrixRemoveMatComboList(MDC_fCombo As SAPbouiCOM.Column)
    Dim i As Long
    For i = 1 To MDC_fCombo.ValidValues.Count
        Call MDC_fCombo.ValidValues.Remove(0, psk_Index)
    Next
End Sub

'--------------------------------------------------------------------------------------
'//     메세지 출력 펑션
'--------------------------------------------------------------------------------------
Public Function MDC_GF_Message(MDC_pMsg As String, MDC_pType As String) As Long    '//정상메세지
    
    Select Case UCase(MDC_pType)
    Case "S"
        Sbo_Application.StatusBar.SetText MDC_pMsg, bmt_Short, smt_Success
    Case "E"
        Sbo_Application.StatusBar.SetText MDC_pMsg, bmt_Short, smt_Error
    Case "W"
        Sbo_Application.StatusBar.SetText MDC_pMsg, bmt_Short, smt_Warning
    End Select
    
End Function


Public Function GovIDCheck(sID As String) As Boolean
'*******************************************************
' 주민등록번호/외국인등록번호 오류체크
' Argument  : sID : 검증하려는 번호
' Remark    : 외국인 주민번호 : 000000-1234567(앞의 6자리는 생년월일,뒷자리는 1:성별구분, 23:등록기관번호, 45:일련번호, 6:등록자구분, 7:검증번호)
'********************************************************
    Dim Weight  As String
    Dim Total   As Integer
    Dim chk     As Integer
    Dim Rmn     As Integer
    Dim i       As Integer
    Dim Dt      As Integer
    Dim Wt      As Integer
    
    GovIDCheck = False
    
    sID = Trim$(sID)
    If sID = "" Then Exit Function
    If Mid$(sID, 7, 1) = "-" Then sID = Left$(sID, 6) & Mid$(sID, 8)
    If Len(sID) <> 13 Then Exit Function
    
    '// 성별구분코드(1,2,3,4:내국인, 5,6,7,8:외국인)
    If Mid(sID, 7, 1) < "1" Or Mid(sID, 7, 1) > "8" Then Exit Function
    '// 검증코드
    Select Case Mid(sID, 7, 1)
    Case "5", "6", "7", "8"   '// 외국인
        If (Val(Mid(sID, 8, 2)) Mod 2) <> 0 Then   '// 등록기관번호검증
            Exit Function
        End If
    End Select
    
    chk = Val(Right$(sID, 1))

    Total = 0
    Weight = "234567892345"
    
    For i = 1 To 12
        Dt = Val(Mid$(sID, i, 1))
        Wt = Val(Mid$(Weight, i, 1))
        Total = Total + (Dt * Wt)
    Next i

    Rmn = 11 - (Total Mod 11)
    
    If Rmn > 9 Then Rmn = Rmn Mod 10
    
    Select Case Mid(sID, 7, 1)
    Case "5", "6", "7", "8"   '// 외국인
        Rmn = Rmn + 2
        If Rmn >= 10 Then Rmn = Rmn - 10
    End Select
    
    GovIDCheck = IIf(Rmn = chk, True, False)
End Function

Function Age_Chk(Str As String, StdDate$) As PerSonNo_Info
    '***************************************************************************
    'Function ID : 주민번호체크 나이구함.
    '기    능    :
    '인    수    : 수치
    '반 환 값    : None
    '특이사항    : None
    '추가        : 함미경
    '***************************************************************************
    Dim chk As PerSonNo_Info
    Dim PerNo As String
    Dim Bir_YY As String
    
    PerNo = Replace(Str, "-", "")  '/ 주민등록번호
    With chk
        .check = True
        Select Case Mid$(PerNo, 7, 1)
        Case "1":        Bir_YY = "19": .Sex = 0
        Case "2":        Bir_YY = "19": .Sex = 1   '/ 여자
        Case "5":        Bir_YY = "19": .Sex = 0
        Case "6":        Bir_YY = "19": .Sex = 1
        Case "3":        Bir_YY = "20": .Sex = 0
        Case "4":        Bir_YY = "20": .Sex = 1
        Case "7":        Bir_YY = "20": .Sex = 0
        Case "8":        Bir_YY = "20": .Sex = 1
        Case "9":        Bir_YY = "18": .Sex = 0
        Case "0":        Bir_YY = "18": .Sex = 1
        End Select
        .BirthDay = Format$(Trim$(Bir_YY) & Left$(PerNo, 6), "0000-00-00")   '/ 생일
        If Len(PerNo) <> 13 Then .check = False
        '일자체크
        If IsDate(.BirthDay) = False Then    '/ 날짜 체크
            .check = False
            .ManAge = 0
            .Age = 0
        Else
            .ManAge = DateDiff("yyyy", .BirthDay, Format$(StdDate, "0000-00-00"))
            .Age = .ManAge + 1
        End If
    End With
    
    Age_Chk = chk
End Function
