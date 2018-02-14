Attribute VB_Name = "MDC_Com"
'*******************************************************************************
' 화면  ID : ComControl
' 화 면 명 :
' 기    능 : 클래스 모듈
' Table 명 : None
' 입력  값 :
' 출력  값 :
' 작 성 자 :
' 작 성 일 : 2002. 11. 25
'-------------------------------------------------------------------------------
' 수 정 일    |    수 정 자    |                   수 정 내 용
'-------------------------------------------------------------------------------
'
'*******************************************************************************
Option Explicit

Public Function GetSpStr(Str As String) As String
    '***************************************************************************
    'Function ID : GetSpStr
    '기    능    : 한칸 뛰워진 값의 앞에값을 가져온다
    '인    수    : Str
    '반 환 값    : None
    '특이사항    : ex) Str = 0 강대봉, 리턴 = 강대봉
    '***************************************************************************
    Dim i As Long
   
    i = InStr(1, Str, " ", vbTextCompare)
    
    If i > 0 Then
        GetSpStr = Mid(Str, 1, i - 1)
    Else
        GetSpStr = ""
    End If
End Function

Public Function GetSpStr2(Str As String) As String
    '***************************************************************************
    'Function ID : GetSpStr2
    '기    능    : 한칸 뛰워진 값의 뒤에값을 가져온다
    '인    수    : Str
    '반 환 값    : None
    '특이사항    : ex) Str = 0 강대봉, 리턴 = 0
    '***************************************************************************
    Dim Buf As String
    
    Buf = GetSpStr(Str)

    If Len(Str) > Len(Buf) + 1 Then
        GetSpStr2 = Mid(Str, Len(Buf) + 2, Len(Str) - Len(Buf) - 1)
    End If
End Function

Public Function uISDATE(Dt As Variant, Conv As Variant) As Variant
    '***************************************************************************
    'Function ID : uISDATE
    '기    능    : 날짜형식 가부의 판정
    '인    수    : dt,Conv
    '반 환 값    : None
    '특이사항    : None
    '***************************************************************************
    If IsDate(Dt) Then
        uISDATE = Dt
    Else
        uISDATE = Conv
    End If
End Function

Public Function uISNULL(Str As Variant, Conv As Variant) As Variant
    '***************************************************************************
    'Function ID : uISNULL
    '기    능    : 문자형식 가부의 판정
    '인    수    : Str,Conv
    '반 환 값    : None
    '특이사항    : None
    '***************************************************************************
    uISNULL = IIf(IsNull(Str), Conv, Str)
End Function

Public Function uISNUMERIC(num As Variant, Conv As Variant, NumType As String) As Variant
    '***************************************************************************
    'Function ID : uISNUMERIC
    '기    능    : 숫자형식의 변환
    '인    수    : Str,Conv
    '반 환 값    : None
    '특이사항    : None
    '***************************************************************************
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
    '***************************************************************************
    'Function ID : GetLeftNumZero
    '기    능    : 숫자를 오른정렬하면서 왼쪽엔 '0'를 채우는 함수
    '인    수    : Num,자리수
    '반 환 값    : None
    '특이사항    :
    '***************************************************************************
    ' ▶ex) GetLeftNumZero("123456",10) --> 반환값 = "0000123456"
    '***************************************************************************
    Dim Ln As Long
    
    Ln = Len(num)
    If Ln <= 자리수 Then
        GetLeftNumZero = String(자리수 - Ln, "0") + num
    Else
        GetLeftNumZero = String(자리수, "0")
    End If
End Function



''------------------
''/ 환율을 리턴
''------------------
'Public Function GetExchangRate(ByVal pCurrency$, ByVal pDate As String) As Double
'GoTo GetExchangRate_Error
'    Dim sSQL            As String
'    Dim oRecordSet      As SAPbobsCOM.Recordset
'
'    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
'
'    sSQL = ""
'    sSQL = sSQL & "  SELECT  Rate  "
'    sSQL = sSQL & "    FROM  [ORTT] "
'    sSQL = sSQL & "   WHERE  Currency = '" & pCurrency & "' "
'    sSQL = sSQL & "          AND RateDate = '" & pDate & "' "
'
'    oRecordSet.DoQuery sSQL
'
'    If oRecordSet.RecordCount > 0 Then
'        GetExchangRate = oRecordSet.fields(0).Value
'    Else
'        GetExchangRate = -1
'    End If
'
'    Set oRecordSet = Nothing
'    Exit Function
'GetExchangRate_Error:
'    GetExchangRate = -1
'    Set oRecordSet = Nothing
'End Function



'Function HanAmt(Money As String) As String
'
'    Dim Ln As Integer
'    Dim Mon1 As Integer, Mon2 As Integer, Mon3 As Integer
'    'Dim 한자1 As Variant, 한자2 As Variant
'
'    Dim RetMeney    As String
'    Dim arrHanAmt   As String
'    Dim arrUnit1    As Integer
'    Dim arrUnit2    As String
'    Dim i, j          As Integer
'
'    Dim Str$
'
'    '한자1 = Array("壹", "貳", "參", "四", "五", "六", "七", "八", "九", "拾")
'    '한자2 = Array("壹", "拾", "百", "阡", "萬", "拾", "百", "阡", "億", "拾", "百", "阡", "兆", "拾", "百", "阡")
'
'    '1,001,001,001,000
'    '1조,111,1억11,11만1천,111
'
'
'    arrHanNum = Array("일", "이", "삼", "사", "오", "육", "칠", "팔", "구", "십")
'    arrUnit1 = Array("십", "백", "천")
'    arrUnit2 = Array("만", "억", "조")
'
'    Ln = Len(Money)
'
'    Money = Format(CCur(Money), "0")
'
'    For i = 1 To 9
'        Money = Replace(Money, 1, arrHanNum(i))
'    Next
'
'
'    Select Case Ln
'    Case (5 > Ln)
'        For i = 1 To 3
'            If Ln < 2 Then
'                Exit For
'            End If
'
'            Select Case Ln
'            Case 2
'                RetMeney = Left(Money, 1) & arrUnit1(i)
'            Case 3
'                RetMeney = Left(Money, 1) & arrUnit1(i)
'            Case 4
'                RetMeney = Left(Money, 1) & arrUnit1(i)
'            End Select
'
'            Money = Mid(Money, 2, Ln)
'            Ln = Ln - 1
'        Next
'
'    Case (5 <= Ln) And (Ln < 9)
'
'    Case (9 <= Ln) And (Ln < 13)
'
'    End Select
'
'
'    If Money <> 0 Then
'    Ln = Len(Money)
'    If Ln > 12 Then
'        If Mid(Money, Ln - 12, 1) = "0" Then Mon1 = 1
'    End If
'
'    If Ln > 8 Then
'        If Mid(Money, Ln - 8, 1) = "0" Then Mon2 = 1
'        If CInt(Mid(Money, Ln - 7, 4)) > 0 And (Mid(Money, Ln - 4, 1) = "0") Then Mon3 = 1
'    ElseIf Ln > 4 And (Mid(Money, Ln - 4, 1) = "0") Then
'        Mon3 = 1
'    End If
'
'    For i = 1 To Ln
'        If Mid(Money, i, 1) <> "0" Then
'            Str = Str + 한자1(CInt(Mid(Money, i, 1)) - 1)
'            Str = Str + 한자2(Ln - i)
'        Else
'            If ((Ln - i + 1) = 5 And Mon3 > 0) Or ((Ln - i + 1) = 9 And Mon2 > 0) Or ((Ln - i + 1) = 13 And Mon1 > 0) Then
'                 Str = Str + 한자2(Ln - i)
'            End If
'
'        End If
'    Next i
'    End If
'
'    Erase arrHanAmt
'    Erase arrUnit
'
'    HanAmt = Str
'End Function


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

'---------------------------------------------------------------------------------------
''//    CHOOSEFROMLIST의 값을 리턴
'       MDC_GP_ChooseFromList_DBDatasourceReturn(PVAL, FORMUID, 테이블이름, 리턴할 컬럼,
'       MATRIX, 시작 ROW, 라인번호컬럼, 체크박스일 경우 컬럼명, 체크박스 초기값)
'---------------------------------------------------------------------------------------
Public Sub MDC_GP_CF_DBDatasourceReturn(pval As SAPbouiCOM.IItemEvent, _
                                                      MDC_pFormUID As String, _
                                                      MDC_pTableName As String, _
                                                      Optional ByVal MDC_sUDS As String = "", _
                                                      Optional ByVal MDC_pMatrix As String = "", _
                                                      Optional ByVal MDC_pRow As Integer = 0, _
                                                      Optional ByVal MDC_pSeqNoUDS As String = "", _
                                                      Optional ByVal MDC_pFieldName As String = "", _
                                                      Optional ByVal MDC_pFieldValue As String = "")

    Dim MDC_oCFLEvento  As SAPbouiCOM.IChooseFromListEvent
    Dim MDC_sCFLID      As String
    Dim MDC_oCFL        As SAPbouiCOM.ChooseFromList
    Dim MDC_oDataTable  As SAPbouiCOM.DataTable
    
    Dim MDC_pForm       As SAPbouiCOM.Form
    Dim MDC_oMatrix     As SAPbouiCOM.Matrix
    Dim MDC_oDBTable    As SAPbouiCOM.DBDataSource
    
    Dim MDC_iLooper     As Integer
    Dim MDC_jLooper     As Integer
    Dim MDC_sTemp01
    
    Set MDC_pForm = Sbo_Application.Forms.Item(MDC_pFormUID)

    Set MDC_oCFLEvento = pval
    Set MDC_oDataTable = MDC_oCFLEvento.SelectedObjects
    MDC_sCFLID = MDC_oCFLEvento.ChooseFromListUID
    '// 취소버튼 클릭시
    If MDC_oDataTable Is Nothing Then
        Exit Sub
    End If
    
    Set MDC_oCFL = MDC_pForm.ChooseFromLists.Item(MDC_sCFLID)
    Set MDC_oDBTable = MDC_pForm.DataSources.DBDataSources.Item(MDC_pTableName)
    If MDC_pMatrix <> "" Then Set MDC_oMatrix = MDC_pForm.Items(MDC_pMatrix).Specific
    MDC_sTemp01 = Split(MDC_sUDS, ",")
    
    If MDC_pMatrix <> "" And MDC_pRow > 0 Then
    
        For MDC_jLooper = 0 To MDC_oDataTable.Rows.Count - 1
            
            If MDC_jLooper > 0 Then
                If MDC_pSeqNoUDS <> "" Then
                    MDC_oDBTable.InsertRecord (MDC_pRow + MDC_jLooper - 1)
                    MDC_oDBTable.Offset = MDC_pRow + MDC_jLooper - 1
                    MDC_oDBTable.setValue MDC_pSeqNoUDS, MDC_pRow + MDC_jLooper - 1, MDC_pRow + MDC_jLooper
                Else
                    MDC_oDBTable.InsertRecord (MDC_pRow + MDC_jLooper - 1)
                    MDC_oDBTable.Offset = MDC_pRow + MDC_jLooper - 1
                End If
            Else
                MDC_oDBTable.Offset = MDC_pRow + MDC_jLooper - 1
            End If
            
            For MDC_iLooper = 0 To UBound(MDC_sTemp01)
                If MDC_oCFL.ObjectType = "171" Then   '// 사원마스타일경우 성 + 이름
                    If MDC_iLooper = 1 Then
                        MDC_oDBTable.setValue MDC_sTemp01(MDC_iLooper), MDC_pRow + MDC_jLooper - 1, MDC_oDataTable.GetValue(MDC_iLooper, MDC_jLooper) + MDC_oDataTable.GetValue(MDC_iLooper + 1, MDC_jLooper)
                    Else
                        MDC_oDBTable.setValue MDC_sTemp01(MDC_iLooper), MDC_pRow + MDC_jLooper - 1, MDC_oDataTable.GetValue(MDC_iLooper, MDC_jLooper)
                    End If
                Else
                    MDC_oDBTable.setValue MDC_sTemp01(MDC_iLooper), MDC_pRow + MDC_jLooper - 1, MDC_oDataTable.GetValue(MDC_iLooper, MDC_jLooper)
                End If
            Next MDC_iLooper
            
            If MDC_pFieldName <> "" And MDC_pFieldValue <> "" Then MDC_oDBTable.setValue MDC_pFieldName, MDC_pRow + MDC_jLooper - 1, MDC_pFieldValue
            
            MDC_oMatrix.LoadFromDataSource
        Next MDC_jLooper
    Else
        For MDC_iLooper = 0 To UBound(MDC_sTemp01)
        
            Select Case MDC_oCFL.ObjectType
            Case "171"            '// 사원마스타
                If MDC_iLooper = 1 Then
                    MDC_oDBTable.setValue MDC_sTemp01(MDC_iLooper), 0, MDC_oDataTable.GetValue(MDC_iLooper, 0) + MDC_oDataTable.GetValue(MDC_iLooper + 1, 0)
                Else
                    MDC_oDBTable.setValue MDC_sTemp01(MDC_iLooper), 0, MDC_oDataTable.GetValue(MDC_iLooper, 0)
                End If
            Case "17", "22"    '// 판매오더, 생산오더, 구매오더
                MDC_oDBTable.setValue MDC_sTemp01(MDC_iLooper), 0, MDC_oDataTable.GetValue(MDC_iLooper + 1, 0)
            Case "202"
                MDC_oDBTable.setValue MDC_sTemp01(MDC_iLooper), 0, MDC_oDataTable.GetValue(MDC_iLooper + 3, 0)
            Case "CPG001"
                If MDC_iLooper = 0 Then
                    MDC_oDBTable.setValue MDC_sTemp01(MDC_iLooper), 0, MDC_oDataTable.GetValue("U_PRJCODE", 0)
                Else
                    MDC_oDBTable.setValue MDC_sTemp01(MDC_iLooper), 0, MDC_oDataTable.GetValue("U_PRJNAME", 0)
                End If
            Case Else
                MDC_oDBTable.setValue MDC_sTemp01(MDC_iLooper), 0, MDC_oDataTable.GetValue(MDC_iLooper, 0)
            End Select
            
        Next MDC_iLooper
    End If
    
    Exit Sub

End Sub
'// Matrix Combo Box Setting
Public Sub MDC_GP_MatrixSetMatComboList(MDC_fCombo As SAPbouiCOM.Column, _
                                        MDC_fSQL As String, _
                                        Optional AndLine$, _
                                        Optional AddSpace$)
    '***************************************************************************
    'Function ID : GetListIndex
    '기    능    :
    '인    수    : Lst
    '반 환 값    : None
    '특이사항    : 콤보박스의 들어가야 할 내용을 시스템 코드에서 가져와 세팅한다
    '***************************************************************************
    Dim MDC_fRecordset As SAPbobsCOM.Recordset
    
    Set MDC_fRecordset = Sbo_Company.GetBusinessObject(BoRecordset)
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
Public Sub MDC_GP_RowSelect_Delete(MDC_dForm As SAPbouiCOM.Form, _
                            MDC_dMatrix As SAPbouiCOM.Matrix, _
                            MDC_iRow As Integer, _
                            MDC_cColumn As String)

    Dim MDC_iLooper         As Integer
    
    MDC_dMatrix.DeleteRow MDC_iRow
    
    For MDC_iLooper = 1 To MDC_dMatrix.VisualRowCount
        MDC_dMatrix.Columns(MDC_cColumn).Cells(MDC_iLooper).Specific.VALUE = MDC_iLooper
    Next MDC_iLooper
    
    MDC_dMatrix.FlushToDataSource
    MDC_dMatrix.Clear
    MDC_dMatrix.LoadFromDataSource
    
    If MDC_dForm.Mode = fm_OK_MODE Then MDC_dForm.Mode = fm_UPDATE_MODE
    
    Set MDC_dMatrix = Nothing
    
End Sub

'--------------------------------------------------------------------------------------
'//     NULL 값 체크
'--------------------------------------------------------------------------------------
Public Function MDC_GF_Nz(MDC_pAnyData) As Currency
    
    On Error GoTo Err_Disp
    
    If MDC_pAnyData = "" Then MDC_pAnyData = 0
    
    If Not IsNumeric(MDC_pAnyData) Then MDC_pAnyData = 0
    
    MDC_GF_Nz = IIf(IsNull(MDC_pAnyData), 0, MDC_pAnyData)
    
Exit Function

Err_Disp:
    
    MDC_pAnyData = 0
    
End Function
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


