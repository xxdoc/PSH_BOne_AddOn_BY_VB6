Attribute VB_Name = "MDC_GetData"
Option Explicit

Public Function Get_ReData(oReColumn$, oColumn$, oTable$, oTaValue$, Optional AndLine$) As Variant
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '반환컬럼,조건 컬럼,테이블,조건값,앤드절
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String

    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    
    sQry = "SELECT " & oReColumn & " FROM " & oTable
    sQry = sQry & " WHERE " & oColumn & " = " & oTaValue
    If AndLine <> "" Then
        sQry = sQry & AndLine
    End If
    oRecordSet.DoQuery sQry

    Do Until oRecordSet.EOF
        Get_ReData = oRecordSet(0).Value
        oRecordSet.MoveNext
    Loop

    Set oRecordSet = Nothing
End Function


Public Function Get_Series_No(ObjectCode$) As Long
    Dim f_RecordSet As SAPbobsCOM.Recordset
    Dim Sql$
    Dim Price_List$

    Set f_RecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Sql = "SELECT Series  FROM nnm1"
    Sql = Sql + " WHERE ObjectCode='" + ObjectCode + "'"
    f_RecordSet.DoQuery Sql
    
    If f_RecordSet.RecordCount > 0 Then
        Get_Series_No = Trim(f_RecordSet.Fields(0).Value)
    Else
        Get_Series_No = ""
    End If
    
    Set f_RecordSet = Nothing
    
End Function

' DataType = db_Date 일 경우에 운영관리의 Separator가 분리자를 불러와
' 날짜 형식을 "9999-99-99"로 변경을 해 줘야 한다.
Public Function GP_DateSeparatorChange(ByVal pDate As String, Optional pTrue As Boolean = True) As String

    If IsNull(pDate) Then
        GP_DateSeparatorChange = ""
        Exit Function
    End If
    
    If Len(pDate) = 8 Then
        GP_DateSeparatorChange = Format$(Left$(pDate, 4) & "-" & Mid$(pDate, 5, 2) & "-" & Mid$(pDate, 7, 2), "YYYY-MM-DD")
        Exit Function
    End If
    
    pDate = Replace(Replace(pDate, ".", "-"), "/", "-")
    
    If pTrue = True Then
        GP_DateSeparatorChange = Format$(pDate, "YYYY-MM-DD")
    Else
        GP_DateSeparatorChange = Format$(pDate, "YYYYMMDD")
    End If
    
    Exit Function
    
End Function

