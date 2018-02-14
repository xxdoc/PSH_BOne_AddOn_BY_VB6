Attribute VB_Name = "MDC_Company_Common"
'//ODBC연결
Public Function ConnectODBC() As Boolean
    On Error Resume Next
    
    ConnectODBC = False
        
    
    '//ODBC방식
    ZG_CRWDSN = "PROVIDER=MSDASQL;DSN=" & MDC_Globals.SP_ODBC_Name & ";DATABASE=" & oCompany.CompanyDB & ";UID=" & MDC_Globals.SP_ODBC_ID & ";PWD=" & MDC_Globals.SP_ODBC_PW & ";"
    
    Set g_ERPDMS = New ADODB.Connection
    g_ERPDMS.ConnectionTimeout = 30
    g_ERPDMS.CursorLocation = adUseClient
    g_ERPDMS.Open ZG_CRWDSN
    If Err <> 0 Then
        Sbo_Application.SetStatusBarMessage "ODBC데이터베이스 연결에 실패하였습니다. ODBC설정을 확인하십시오!! ", bmt_Short, False
    Else
        ConnectODBC = True
    End If

'        '//SQLOLEDB방식
'        ZG_CRWDSN = "Provider=SQLOLEDB.1;Server=" & oCompany.Server & ",1433;uid=" & MDC_Globals.SP_ODBC_ID & ";pwd=" & MDC_Globals.SP_ODBC_PW & ";database=" & oCompany.CompanyDB & ";Connect Timeout=180"
'        Set g_ERPDMS = New ADODB.Connection
'        g_ERPDMS.ConnectionTimeout = 30
'        g_ERPDMS.CursorLocation = adUseClient
'        g_ERPDMS.Open ZG_CRWDSN
'        If Err <> 0 Then
'            Sbo_Application.SetStatusBarMessage "SQLOLDDB 연결에 실패하였습니다. " & Err & "코드 ODBC설정을 확인하십시오!! ", bmt_Short, False
'        Else
'            ConnectODBC = True
'        End If

    
End Function

'//쿼리실행
'//DoQuery("쿼리변수")
Public Sub DoQuery(ByVal sQry As String)
    Dim oRecordset As SAPbobsCOM.Recordset
    Set oRecordset = oCompany.GetBusinessObject(BoRecordset)
    Call oRecordset.DoQuery(sQry)
    Set oRecordset = Nothing
End Sub

'//쿼리실행
'//GetValue("쿼리변수","필드위치","레코드위치")
Public Function GetValue(ByVal sQry As String, Optional ByVal FieldCount As Long, Optional ByVal RecordCount As Long) As Variant
    Dim i As Long
    Dim oRecordset As SAPbobsCOM.Recordset
    Set oRecordset = oCompany.GetBusinessObject(BoRecordset)
    Call oRecordset.DoQuery(sQry)
    If (oRecordset.RecordCount > 0) Then
        oRecordset.MoveFirst
        If (RecordCount = 0) Then
            RecordCount = 1
        End If
        For i = 1 To RecordCount
            GetValue = oRecordset.Fields(FieldCount).Value
            oRecordset.MoveNext
        Next
    Else
        GetValue = ""
    End If
    Set oRecordset = Nothing
End Function

Public Sub ActiveUserDefineValue(ByRef oForm As SAPbouiCOM.Form, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean, ByVal ItemUID As String, Optional ByVal ColumnUID As String)
    If ColumnUID = "" Then
        If pval.ItemUID = ItemUID Then
            If pval.CharPressed = "9" Then
                If oForm.Items(ItemUID).Specific.Value = "" Then
                    Sbo_Application.ActivateMenuItem ("7425")
                    BubbleEvent = False
                End If
            End If
        End If
    Else
        If pval.ItemUID = ItemUID Then
            If pval.ColUID = ColumnUID Then
                If pval.CharPressed = "9" Then
                    If oForm.Items(ItemUID).Specific.Columns(ColumnUID).Cells(pval.Row).Specific.Value = "" Then
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                End If
            End If
        End If
    End If
End Sub

Public Sub ActiveUserDefineValueAlways(ByRef oForm As SAPbouiCOM.Form, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean, ByVal ItemUID As String, Optional ByVal ColumnUID As String)
    If ColumnUID = "" Then
        If pval.ItemUID = ItemUID Then
            If pval.CharPressed = "9" Then
                If oForm.Items(ItemUID).Specific.Value = "" Then
                    Sbo_Application.ActivateMenuItem ("7425")
                    BubbleEvent = False
                End If
            Else
                Sbo_Application.ActivateMenuItem ("7425")
                BubbleEvent = False
            End If
        End If
    Else
        If pval.ItemUID = ItemUID Then
            If pval.ColUID = ColumnUID Then
                If pval.CharPressed = "9" Then
                    If oForm.Items(ItemUID).Specific.Columns(ColumnUID).Cells(pval.Row).Specific.Value = "" Then
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                Else
                    Sbo_Application.ActivateMenuItem ("7425")
                    BubbleEvent = False
                End If
            End If
        End If
    End If
End Sub

