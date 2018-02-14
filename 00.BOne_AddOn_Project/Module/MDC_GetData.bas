Attribute VB_Name = "MDC_GetData"
'*******************************************************************************
' 화면  ID : MDC_GetData
' 화 면 명 :
' 기    능 : 모듈(원가처리시에 정보를 반환하는 함수의 집합)
' Table 명 : None
' 입력  값 :
' 출력  값 :
' 작 성 자 : 김영호,함미경
' 작 성 일 : 2005. 08. 22~~~~~~~~~~~~~~~~~~
'//  Copyright  (c) Morning Data
'-------------------------------------------------------------------------------
' 수 정 일    |    수 정 자    |                   수 정 내 용
'-------------------------------------------------------------------------------
Option Explicit

Public Sub SpaceLineDelete(TableId$, KeyColumn$)
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '테이블에 비어있는 라인을 삭제한다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry$
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    sQry = "Delete " & TableId
    sQry = sQry + " WHERE ISNULL(" & KeyColumn & ",'')=''"
    oRecordSet.DoQuery sQry
    
    Set oRecordSet = Nothing
End Sub

Public Sub SetMessage_Err(Msg As String)  '에러메세지 출력
    Call Sbo_Application.SetStatusBarMessage(Msg, bmt_Short, True)
End Sub

Public Sub SetMessage_Basic(Msg As String)   '성공메세지 출력
    Call Sbo_Application.SetStatusBarMessage(Msg, bmt_Short, False)
End Sub

Public Function Get_ItemName(ItemCode$) As String
'// 품명을 반환합니다
    Dim oRecordSet As SAPbobsCOM.Recordset
    Dim SQL        As String
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    SQL = "select ItemName from OITM WHERE ItemCode='" & ItemCode & "'"
    oRecordSet.DoQuery SQL
    Do Until oRecordSet.EOF
        Get_ItemName = oRecordSet.Fields(0).VALUE      '/
        oRecordSet.MoveNext
    Loop
    
    Set oRecordSet = Nothing
    Exit Function
'//////////////////////////////////////////////////////////////////////////////////////
LenDecimal_Error:
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "품명을 가져올수 없습니다." & Space(10) & Err.Description, bmt_Short, True
End Function

Public Sub Set_BPLIdIndex(ByRef G_FormSet As SAPbouiCOM.Form, ComboUid$)
    Dim oRecordSet  As SAPbobsCOM.Recordset
    Dim sQry$
    Dim oCombobox As SAPbouiCOM.ComboBox
    
    Set oCombobox = G_FormSet.Items(ComboUid).Specific
    
    If oCombobox.ValidValues.Count > 0 Then
        Call oCombobox.Select(0, psk_Index)
    End If
    
    Set oCombobox = Nothing
End Sub

Public Function Get_ReData(oReColumn$, oColumn$, oTable$, oTaValue$, Optional AndLine$) As Variant
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '반환컬럼,조건 컬럼,테이블,조건값,앤드절
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String

    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)

    sQry = "SELECT " & oReColumn & " FROM " & oTable
    sQry = sQry & " WHERE " & oColumn & " = " & oTaValue
    If AndLine <> "" Then
        sQry = sQry & AndLine
    End If
    oRecordSet.DoQuery sQry

    Do Until oRecordSet.EOF
        Get_ReData = oRecordSet(0).VALUE
        oRecordSet.MoveNext
    Loop

    Set oRecordSet = Nothing
End Function

'///////////////////////////////
'------------------------------
'/ 생산오더의 DocEntry
'------------------------------
'//////////////////////////////
Public Function GetOWOR_DocEntry(ByVal DocNum$) As String
    Dim sSQL         As String
    Dim oRecordSet   As SAPbobsCOM.Recordset

    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)

    sSQL = ""
    sSQL = sSQL & " SELECT  DocEntry "
    sSQL = sSQL & " FROM    [OWOR]  "
    sSQL = sSQL & " WHERE   DocNum = '" & DocNum & "' "

    oRecordSet.DoQuery sSQL
    
    GetOWOR_DocEntry = oRecordSet.Fields("DocEntry").VALUE
    
    If Trim(GetOWOR_DocEntry) = "" Then
        GetOWOR_DocEntry = "F"
    End If
    
    Set oRecordSet = Nothing
End Function

'------------------------------------------
'/ 테이블이나 프러시저가 있는지 체크한다.
'-----------------------------------------
Public Function ObjectExitsChk(ByVal OBJECT_ID$) As Boolean
On Error GoTo ObjectExitsChk_Error:
    
    Dim sSQL
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    ObjectExitsChk = False
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    sSQL = ""
    sSQL = sSQL & "  SELECT ISNULL(Object_ID('" & OBJECT_ID & "' ),-9999 )     "
    
    oRecordSet.DoQuery sSQL

    If oRecordSet.Fields(0).VALUE = -9999 Then
        ObjectExitsChk = False
        Set oRecordSet = Nothing
        Exit Function
    End If

    Set oRecordSet = Nothing
    
    ObjectExitsChk = True

    Exit Function
ObjectExitsChk_Error:
    Set oRecordSet = Nothing
    ObjectExitsChk = False
End Function


'------------------------------------------
'/ 해당 클레스의 테이블등록 체크
'-----------------------------------------
Public Function ObjectTableChk(ByVal Tablename$) As Boolean
On Error GoTo ObjectTableChk_Error:
    
    ObjectTableChk = True

    If ObjectExitsChk(Tablename) = False Then
        If Z_Language = 28 Then
            Sbo_Application.SetStatusBarMessage " 테이블  [" & Tablename & "] 이 필요합니다  ! ", bmt_Short, False
        Else
            Sbo_Application.SetStatusBarMessage " Table [" & Tablename & "] is required  ! ", bmt_Short, False
        End If
        ObjectTableChk = False
        Exit Function
    End If
    
    ObjectTableChk = True
    Exit Function
ObjectTableChk_Error:
    ObjectTableChk = False
End Function

'------------------------------------------
'/ 해당 클레스의 오브젝트등록 체크
'-----------------------------------------
Public Function ClassObjectChk(ByVal Tablename$) As Boolean
On Error GoTo ClassObjectChk_Error:
    Dim sSQL    As String
    Dim oRecordSet As SAPbobsCOM.Recordset

    ClassObjectChk = False

    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)

    sSQL = ""
    sSQL = sSQL & " SELECT  COUNT(*) "
    sSQL = sSQL & " FROM [OUDO]    "
    sSQL = sSQL & " WHERE  TableName  = '" & Tablename & "' "

    oRecordSet.DoQuery sSQL

    If oRecordSet.Fields(0).VALUE <= 0 Then
        If Z_Language = 28 Then
            Sbo_Application.SetStatusBarMessage Tablename & " 의 UDO 등록이 필요합니다  ! ", bmt_Short, False
        Else
            Sbo_Application.SetStatusBarMessage " Registration of  " & Tablename & " is required  ! ", bmt_Short, False
        End If
        
        Set oRecordSet = Nothing
        ClassObjectChk = False
        Exit Function
    End If

    Set oRecordSet = Nothing
    
    ClassObjectChk = True

    Exit Function
ClassObjectChk_Error:
            
    Set oRecordSet = Nothing
    ClassObjectChk = False
    
End Function

'------------------------------------------
'/ 숫자를 한글로 표시
'-----------------------------------------
'Public Function NumberToHanGul(num) As String
'On Error GoTo NumberToHanGul_Error
'     Dim i As Integer, size As Integer, dsize As Integer
'     Dim sw As Boolean
'     Dim tempstring As String
'     Dim temp1 As String, temp2 As String
'     Dim DigitArry1, DigitArry2, NumberArry
'
'     DigitArry1 = Array("", "십", "백", "천")
'     DigitArry2 = Array("", "만", "억", "조")
'     NumberArry = Array("", "일", "이", "삼", "사", "오", "육", "칠", "팔", "구")
'     If Application.IsNumber(num) = False Then
'        'NumberToHanGul = "숫자가 아닙니다."
'        NumberToHanGul = ""
'        Exit Function
'     End If
'     If num = 0 Then Exit Function
'
'     size = Len(num)
'     sw = True
'
'     For i = 1 To size
'       dsize = size - i
'       temp1 = NumberArry(CInt(Mid(num, i, 1)))
'       If temp1 <> "" Then
'          If dsize Mod 4 <> 0 Then
'             temp2 = DigitArry1(dsize Mod 4)
'             sw = True
'          Else
'             temp2 = DigitArry2(dsize \ 4)
'             sw = False
'          End If
'       Else
'          If dsize Mod 4 <> 0 Then
'             temp2 = ""
'          Else
'             If sw = True Then
'                temp2 = DigitArry2(dsize \ 4)
'                sw = False
'             End If
'          End If
'       End If
'       tempstring = tempstring & temp1 & temp2
'     Next i
'     '여기서 필요한 접두어나 접미어를 붙일수 있습니다
'     '  NumberToHanGul = "일금" & tempstring & "원"
'     NumberToHanGul = tempstring
'
'     Exit Function
'
'NumberToHanGul_Error:
'    NumberToHanGul = ""
'    Sbo_Application.SetStatusBarMessage "NumberToHanGul_Error :" & Space(10) & Err.Description, bmt_Short, True
'End Function

'Public Function NumberToHanGul(num) As String
'    Dim i As Integer, size As Integer, dsize As Integer
'    Dim sw As Boolean
'    Dim tempstring As String
'    Dim temp1 As String, temp2 As String
'    Dim 단위배열1, 단위배열2, 숫자배열
'
' 단위배열1 = Array("", "십", "백", "천")
' 단위배열2 = Array("", "만", "억", "조")
' 숫자배열 = Array("", "일", "이", "삼", "사", "오", "육", "칠", "팔", "구")
'
' If Application.IsNumber(num) = False Then
'    NumberToHanGul = "숫자가 아닙니다."
'    Exit Function
' End If
' If num = 0 Then Exit Function
'
' size = Len(num)
' sw = True
'
' For i = 1 To size
'   dsize = size - i
'   temp1 = 숫자배열(CInt(Mid(num, i, 1)))
'   If temp1 <> "" Then
'      If dsize Mod 4 <> 0 Then
'         temp2 = 단위배열1(dsize Mod 4)
'         sw = True
'      Else
'         temp2 = 단위배열2(dsize \ 4)
'         sw = False
'      End If
'   Else
'      If dsize Mod 4 <> 0 Then
'         temp2 = ""
'      Else
'         If sw = True Then
'            temp2 = 단위배열2(dsize \ 4)
'            sw = False
'         End If
'      End If
'   End If
'   tempstring = tempstring & temp1 & temp2
' Next i
' '여기서 필요한 접두어나 접미어를 붙일수 있습니다
' '  NumberToHanGul = "일금" & tempstring & "원"
' NumberToHanGul = tempstring
'
'End Function

