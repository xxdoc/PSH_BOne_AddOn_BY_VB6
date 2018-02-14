Attribute VB_Name = "MDC_PS_Common"
'//������1
'Public Sub ConnectODBC()
'    If (MDC_Globals.ProgramType = "Local") Then
'        gParam_ODBC = "MDCERP"
'        gParam_Server = Sbo_Application.Company.ServerName
'        gParam_DBID = "sa"
'        gParam_DataBase = Sbo_Application.Company.DatabaseName
'        gParam_DBPW = "password1!" '//����
'    ElseIf (MDC_Globals.ProgramType = "Server") Then
'        gParam_ODBC = "MDCERP"
'        gParam_Server = Sbo_Application.Company.ServerName
'        gParam_DBID = "sa"
'        gParam_DataBase = Sbo_Application.Company.DatabaseName
'        gParam_DBPW = "password1!" '//����
'    End If
'    ZG_CRWDSN = "PROVIDER=MSDASQL;DSN=" & gParam_ODBC & ";DATABASE=" & gParam_DataBase & ";UID=" & gParam_DBID & ";PWD=" & gParam_DBPW & ";"
'    On Error Resume Next
'    Set g_ERPDMS = New ADODB.Connection
'    g_ERPDMS.ConnectionTimeout = 30
'    g_ERPDMS.CursorLocation = adUseClient
'    g_ERPDMS.Open ZG_CRWDSN
'    If Err <> 0 Then
'      Sbo_Application.SetStatusBarMessage "ODBC�����ͺ��̽� ���ῡ �����Ͽ����ϴ�. ODBC������ Ȯ���Ͻʽÿ�!! ", bmt_Short, False
'    End If
'End Sub
'//������2
Public Sub ConnectODBC()
    On Error Resume Next
    gParam_ODBC = "MDCERP"
    gParam_Server = Sbo_Application.Company.ServerName
    gParam_DBID = MDC_PS_Common.GetValue("EXEC Profile_SELECT 'SERVERINFO'", 6, 1)
    gParam_DataBase = Sbo_Application.Company.DatabaseName
    gParam_DBPW = MDC_PS_Common.GetValue("EXEC Profile_SELECT 'SERVERINFO'", 7, 1)
    ZG_CRWDSN = "PROVIDER=MSDASQL;DSN=" & gParam_ODBC & ";DATABASE=" & gParam_DataBase & ";UID=" & gParam_DBID & ";PWD=" & gParam_DBPW & ";"
    '//ZG_CRWDSN = "PROVIDER=SQLOLEDB;Data Source=" & gParam_Server & ";Initial Catalog=" & gParam_DataBase & ";User ID=" & gParam_DBID & ";Password=" & gParam_DBPW & ";"
    Set g_ERPDMS = New ADODB.Connection
    g_ERPDMS.ConnectionTimeout = 60
    g_ERPDMS.CommandTimeout = 120
    g_ERPDMS.CursorLocation = adUseClient
    g_ERPDMS.Open ZG_CRWDSN
    If Err <> 0 Then
      Sbo_Application.SetStatusBarMessage "ODBC�����ͺ��̽� ���ῡ �����Ͽ����ϴ�. ODBC������ Ȯ���Ͻʽÿ�!! ", bmt_Short, False
    End If
End Sub

Public Sub Combo_ValidValues_Insert(ByVal FormUID As String, ByVal ItemUID As String, ByVal ColumnUID As String, ByVal VALUE As String, ByVal Description As String)
    MDC_PS_Common.DoQuery ("EXEC COMBO_VALIDVALUES_INSERT '" & FormUID & "','" & ItemUID & "','" & ColumnUID & "','" & VALUE & "','" & Description & "'")
End Sub
Public Sub Combo_ValidValues_SetValueItem(ByRef Combo As SAPbouiCOM.ComboBox, ByVal FormUID As String, ByVal ItemUID As String, Optional ByVal EmptyValue As Boolean)
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    Query01 = "SELECT VALUE,DESCRIPTION FROM COMBO_VALIDVALUES WHERE FORMUID = '" & FormUID & "' AND ITEMUID = '" & ItemUID & "'"
    Call RecordSet01.DoQuery(Query01)
    If (RecordSet01.RecordCount > 0) Then
        For i = 1 To Combo.ValidValues.Count
            Combo.ValidValues.Remove (0)
        Next
        If EmptyValue = True Then
            Call Combo.ValidValues.Add("", "")
        End If
        For i = 1 To RecordSet01.RecordCount
            Call Combo.ValidValues.Add(RecordSet01.Fields(0).VALUE, RecordSet01.Fields(1).VALUE)
            RecordSet01.MoveNext
        Next
    End If
    Set RecordSet01 = Nothing
End Sub

Public Sub Combo_ValidValues_SetValueColumn(ByRef Column As SAPbouiCOM.Column, ByVal FormUID As String, ByVal ItemUID As String, ByVal ColumnUID As String, Optional ByVal EmptyValue As Boolean)
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    Query01 = "SELECT VALUE,DESCRIPTION FROM COMBO_VALIDVALUES WHERE FORMUID = '" & FormUID & "' AND ITEMUID = '" & ItemUID & "' AND COLUMNUID = '" & ColumnUID & "'"
    Call RecordSet01.DoQuery(Query01)
    If (RecordSet01.RecordCount > 0) Then
        For i = 1 To Column.ValidValues.Count
            Call Column.ValidValues.Remove(0, psk_Index)
        Next
        If EmptyValue = True Then
            Call Column.ValidValues.Add("", "")
        End If
        For i = 1 To RecordSet01.RecordCount
            Call Column.ValidValues.Add(RecordSet01.Fields(0).VALUE, RecordSet01.Fields(1).VALUE)
            RecordSet01.MoveNext
        Next
    End If
    Set RecordSet01 = Nothing
End Sub

Public Sub DoQuery(ByVal Query01 As String)
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    Call RecordSet01.DoQuery(Query01)
    Set RecordSet01 = Nothing
End Sub

Public Function GetValue(ByVal Query01 As String, Optional ByVal FieldCount As Long, Optional ByVal RecordCount As Long) As Variant
    Dim i As Long
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    Call RecordSet01.DoQuery(Query01)
    If (RecordSet01.RecordCount > 0) Then
        RecordSet01.MoveFirst
        If (RecordCount = 0) Then
            RecordCount = 1
        End If
        For i = 1 To RecordCount
            GetValue = RecordSet01.Fields(FieldCount).VALUE
            RecordSet01.MoveNext
        Next
    Else
        GetValue = ""
    End If
    Set RecordSet01 = Nothing
End Function

Public Sub ActiveUserDefineValue(ByRef oForm01 As SAPbouiCOM.Form, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean, ByVal ItemUID As String, Optional ByVal ColumnUID As String)
    If ColumnUID = "" Then
        If pval.ItemUID = ItemUID Then
            If pval.CharPressed = "9" Then
                If oForm01.Items(ItemUID).Specific.VALUE = "" Then
                    Sbo_Application.ActivateMenuItem ("7425")
                    BubbleEvent = False
                End If
            End If
        End If
    Else
        If pval.ItemUID = ItemUID Then
            If pval.ColUID = ColumnUID Then
                If pval.CharPressed = "9" Then
                    If oForm01.Items(ItemUID).Specific.Columns(ColumnUID).Cells(pval.Row).Specific.VALUE = "" Then
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                End If
            End If
        End If
    End If
End Sub

Public Sub ActiveUserDefineValueAlways(ByRef oForm01 As SAPbouiCOM.Form, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean, ByVal ItemUID As String, Optional ByVal ColumnUID As String)
    If ColumnUID = "" Then
        If pval.ItemUID = ItemUID Then
            If pval.CharPressed = "9" Then
                If oForm01.Items(ItemUID).Specific.VALUE = "" Then
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
                    If oForm01.Items(ItemUID).Specific.Columns(ColumnUID).Cells(pval.Row).Specific.VALUE = "" Then
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

Public Function GetItem_UnWeight(ByVal ItemCode As String) As String
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    Query01 = "SELECT U_UnWeight FROM [OITM] WHERE ItemCode = '" & ItemCode & "'"
    Call RecordSet01.DoQuery(Query01)
    If (RecordSet01.RecordCount = 0) Then
        GetItem_UnWeight = ""
    Else
        GetItem_UnWeight = RecordSet01.Fields(0).VALUE
    End If
    Set RecordSet01 = Nothing
End Function

Public Function GetItem_ItmBsort(ByVal ItemCode As String) As String
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    Query01 = "SELECT U_ItmBsort FROM [OITM] WHERE ItemCode = '" & ItemCode & "'"
    Call RecordSet01.DoQuery(Query01)
    If (RecordSet01.RecordCount = 0) Then
        GetItem_ItmBsort = ""
    Else
        GetItem_ItmBsort = RecordSet01.Fields(0).VALUE
    End If
    Set RecordSet01 = Nothing
End Function

Public Function GetItem_SbasUnit(ByVal ItemCode As String) As String
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    Query01 = "SELECT U_SBasUnit FROM [OITM] WHERE ItemCode = '" & ItemCode & "'"
    Call RecordSet01.DoQuery(Query01)
    If (RecordSet01.RecordCount = 0) Then
        GetItem_SbasUnit = ""
    Else
        GetItem_SbasUnit = RecordSet01.Fields(0).VALUE
    End If
    Set RecordSet01 = Nothing
End Function

Public Function GetItem_ObasUnit(ByVal ItemCode As String) As String
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    Query01 = "SELECT U_OBasUnit FROM [OITM] WHERE ItemCode = '" & ItemCode & "'"
    Call RecordSet01.DoQuery(Query01)
    If (RecordSet01.RecordCount = 0) Then
        GetItem_ObasUnit = ""
    Else
        GetItem_ObasUnit = RecordSet01.Fields(0).VALUE
    End If
    Set RecordSet01 = Nothing
End Function

Public Function GetItem_Unit1(ByVal ItemCode As String) As String
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    Query01 = "SELECT U_UnitQ1 FROM [OITM] WHERE ItemCode = '" & ItemCode & "'"
    Call RecordSet01.DoQuery(Query01)
    If (RecordSet01.RecordCount = 0) Then
        GetItem_Unit1 = ""
    Else
        GetItem_Unit1 = RecordSet01.Fields(0).VALUE
    End If
    Set RecordSet01 = Nothing
End Function

Public Function GetItem_Spec1(ByVal ItemCode As String) As String
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    Query01 = "SELECT U_Spec1 FROM [OITM] WHERE ItemCode = '" & ItemCode & "'"
    Call RecordSet01.DoQuery(Query01)
    If (RecordSet01.RecordCount = 0) Then
        GetItem_Spec1 = ""
    Else
        GetItem_Spec1 = RecordSet01.Fields(0).VALUE
    End If
    Set RecordSet01 = Nothing
End Function

Public Function GetItem_Spec2(ByVal ItemCode As String) As String
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    Query01 = "SELECT U_Spec2 FROM [OITM] WHERE ItemCode = '" & ItemCode & "'"
    Call RecordSet01.DoQuery(Query01)
    If (RecordSet01.RecordCount = 0) Then
        GetItem_Spec2 = ""
    Else
        GetItem_Spec2 = RecordSet01.Fields(0).VALUE
    End If
    Set RecordSet01 = Nothing
End Function

Public Function GetItem_Spec3(ByVal ItemCode As String) As String
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    Query01 = "SELECT U_Spec3 FROM [OITM] WHERE ItemCode = '" & ItemCode & "'"
    Call RecordSet01.DoQuery(Query01)
    If (RecordSet01.RecordCount = 0) Then
        GetItem_Spec3 = ""
    Else
        GetItem_Spec3 = RecordSet01.Fields(0).VALUE
    End If
    Set RecordSet01 = Nothing
End Function

Public Function GetItem_ManBtchNum(ByVal ItemCode As String) As String
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    Query01 = "SELECT ManBtchNum FROM [OITM] WHERE ItemCode = '" & ItemCode & "'"
    Call RecordSet01.DoQuery(Query01)
    If (RecordSet01.RecordCount = 0) Then
        GetItem_ManBtchNum = ""
    Else
        GetItem_ManBtchNum = RecordSet01.Fields(0).VALUE
    End If
    Set RecordSet01 = Nothing
End Function

Public Function GetItem_TradeType(ByVal ItemCode As String) As String
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    Query01 = "SELECT U_TradeType FROM [OITM] WHERE ItemCode = '" & ItemCode & "'"
    Call RecordSet01.DoQuery(Query01)
    If (RecordSet01.RecordCount = 0) Then
        GetItem_TradeType = ""
    Else
        GetItem_TradeType = RecordSet01.Fields(0).VALUE
    End If
    Set RecordSet01 = Nothing
End Function

Public Sub SBO_SetBackOrderFunction(ByRef oForm01 As SAPbouiCOM.Form)
On Error GoTo SBO_SetBackOrderFunction_Error

    Dim oMat01 As SAPbouiCOM.Matrix
    Set oMat01 = oForm01.Items("38").Specific
    If (oForm01.Mode = fm_OK_MODE) Then
        Exit Sub
    End If
    
    If (oMat01.VisualRowCount > 1) Then
        '//�����۾��� ���߷� - ���� �۾����� ������ �߷��� ������ ����
        Dim i As Long
        Dim Query01 As String
        Dim RecordSet01 As SAPbobsCOM.Recordset
        Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
        Dim BaseType As String
        Dim BaseTable As String
        Dim BaseEntry As Long
        Dim BaseLine As Long
        For i = 1 To oMat01.RowCount - 1
            BaseType = oMat01.Columns("43").Cells(i).Specific.VALUE
            If (BaseType = "-1") Then
                GoTo Continue:
            End If
            BaseEntry = oMat01.Columns("45").Cells(i).Specific.VALUE
            BaseLine = oMat01.Columns("46").Cells(i).Specific.VALUE
            If (BaseType = "17") Then '//�Ǹſ���
                BaseTable = "RDR"
            ElseIf (BaseType = "23") Then '//�ǸŰ���
                BaseTable = "QUT"
            ElseIf (BaseType = "15") Then '//��ǰ
                BaseTable = "DLN"
            ElseIf (BaseType = "16") Then '//�ǸŹ�ǰ
                BaseTable = "RDN"
            ElseIf (BaseType = "13") Then '//AR����
                BaseTable = "INV"
            ElseIf (BaseType = "14") Then '//AR�뺯�޸�
                BaseTable = "RIN"
            ElseIf (BaseType = "22") Then '//���ſ���
                BaseTable = "POR"
            ElseIf (BaseType = "20") Then '//�԰�PO
                BaseTable = "PDN"
            ElseIf (BaseType = "21") Then '//���Ź�ǰ
                BaseTable = "RPD"
            ElseIf (BaseType = "18") Then '//AP����
                BaseTable = "PCH"
            ElseIf (BaseType = "19") Then '//AP�뺯�޸�
                BaseTable = "RPC"
            Else
                Sbo_Application.MessageBox "ȭ��ĸ���� �����ڿ��� ���ǹٶ��ϴ�."
                Exit Sub
            End If
            Query01 = " PS_SBO_GETQUANTITY '" & BaseType & "','" & BaseTable & "','" & BaseEntry & "','" & BaseLine & "'"
            RecordSet01.DoQuery Query01
            oMat01.Columns("U_Qty").Cells(i).Specific.VALUE = Round(RecordSet01.Fields(0).VALUE, 0)
            oMat01.Columns("11").Cells(i).Specific.VALUE = Round(RecordSet01.Fields(1).VALUE, 2)
            oMat01.Columns("1").Cells(oMat01.VisualRowCount).Click ct_Regular
Continue:
        Next
        Set RecordSet01 = Nothing
    End If

    Exit Sub

SBO_SetBackOrderFunction_Error:
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "SBO_SetBackOrderFunction_Error:" & Err.Number & " - " & Err.Description, "E"

End Sub

'// ������ ���ӿ� ���� ����ǥ �߰�
Public Function Make_ItemName(ByVal ItemName$) As String
On Error GoTo Make_ItemName_Error
    Dim i&
    Dim TempItemName$
        
    TempItemName = ""
    For i = 1 To Len(ItemName)
        TempItemName = TempItemName + Mid(ItemName, i, 1)
        If Mid(ItemName, i, 1) = "'" Then
            TempItemName = TempItemName + "'"
        End If
    Next i
    
    Make_ItemName = Trim(TempItemName)
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Make_ItemName_Error:
    TempItemName = ""
    MDC_Com.MDC_GF_Message "User_BPLId_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

'// ���̵� ����� ����
Public Function User_BPLId() As String
On Error GoTo User_BPLId_Error
    Dim sQry            As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    sQry = "Select Branch From [OUSR] Where USER_CODE = '" & Trim(Sbo_Company.UserName) & "'"
    oRecordSet01.DoQuery sQry
    
    User_BPLId = Trim(oRecordSet01.Fields(0).VALUE)
    Set oRecordSet01 = Nothing
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
User_BPLId_Error:
    Set oRecordSet01 = Nothing
    User_BPLId = 0
    MDC_Com.MDC_GF_Message "User_BPLId_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

'// ���̵� â�� ���� [�⺻â�� 1, ���ְ��� 8, �Ӱ��� 9]
Public Function User_WhsCode(ByVal Gbn$) As String
On Error GoTo User_WhsCode_Error
    Dim sQry            As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    sQry = "Select a.WhsCode From [OWHS] a Inner Join [OUSR] b On a.BPLid = b.Branch Where b.USER_CODE = '" & Trim(Sbo_Company.UserName) & "' "
    sQry = sQry & "And LEFT(WhsCode, 1) = '" & Gbn & "' And RIGHT(a.WhsCode, 2) = RIGHT(b.DfltsGroup, 2)"
    oRecordSet01.DoQuery sQry
    
    User_WhsCode = Trim(oRecordSet01.Fields(0).VALUE)
    Set oRecordSet01 = Nothing
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
User_WhsCode_Error:
    Set oRecordSet01 = Nothing
    User_WhsCode = 0
    MDC_Com.MDC_GF_Message "User_WhsCode_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

'// ���̵� ��� ����
Public Function User_MSTCOD() As String
On Error GoTo User_MSTCOD_Error
    Dim sQry            As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    sQry = "Select U_MSTCOD From [OHEM] a Inner Join [OUSR] b On a.userId = b.USERID Where b.USER_CODE = '" & Trim(Sbo_Company.UserName) & "'"
    oRecordSet01.DoQuery sQry
    
    User_MSTCOD = Trim(oRecordSet01.Fields(0).VALUE)
    
    Set oRecordSet01 = Nothing
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
User_MSTCOD_Error:
    Set oRecordSet01 = Nothing
    User_MSTCOD = 0
    MDC_Com.MDC_GF_Message "User_MSTCOD_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

'// ���̵� �μ� ����
Public Function User_DeptCode() As String
On Error GoTo User_DeptCode_Error
    Dim sQry            As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    sQry = "Select dept From [OHEM] a Inner Join [OUSR] b On a.userId = b.USERID Where USER_CODE = '" & Trim(Sbo_Company.UserName) & "'"
    oRecordSet01.DoQuery sQry
    
    User_DeptCode = Trim(oRecordSet01.Fields(0).VALUE)
    Set oRecordSet01 = Nothing
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
User_DeptCode_Error:
    Set oRecordSet01 = Nothing
    User_DeptCode = 0
    MDC_Com.MDC_GF_Message "User_DeptCode_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Public Function User_TeamCode() As String
'******************************************************************************
'Function ID : User_TeamCode()
'�ش���    : MDC_PS_Common
'��    ��    : ������ ������� ���ڵ� ��ȸ
'��    ��    : ����
'�� ȯ ��    : TeamCode
'Ư�̻���    : ����
'******************************************************************************
On Error GoTo User_TeamCode_Error

    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    sQry = "Select U_TeamCode From [OHEM] a Inner Join [OUSR] b On a.userId = b.USERID Where USER_CODE = '" & Trim(Sbo_Company.UserName) & "'"
    oRecordSet01.DoQuery sQry
    
    User_TeamCode = Trim(oRecordSet01.Fields(0).VALUE)
    Set oRecordSet01 = Nothing
    Exit Function

User_TeamCode_Error:
    Set oRecordSet01 = Nothing
    User_TeamCode = 0
    MDC_Com.MDC_GF_Message "User_TeamCode_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Public Function User_RspCode() As String
'******************************************************************************
'Function ID : User_RspCode()
'�ش���    : MDC_PS_Common
'��    ��    : ������ ������� ����ڵ� ��ȸ
'��    ��    : ����
'�� ȯ ��    : RspCode
'Ư�̻���    : ����
'******************************************************************************
On Error GoTo User_RspCode_Error

    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    sQry = "Select U_RspCode From [OHEM] a Inner Join [OUSR] b On a.userId = b.USERID Where USER_CODE = '" & Trim(Sbo_Company.UserName) & "'"
    oRecordSet01.DoQuery sQry
    
    User_RspCode = Trim(oRecordSet01.Fields(0).VALUE)
    Set oRecordSet01 = Nothing
    Exit Function

User_RspCode_Error:
    Set oRecordSet01 = Nothing
    User_RspCode = 0
    MDC_Com.MDC_GF_Message "User_RspCode_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Public Function User_ClsCode() As String
'******************************************************************************
'Function ID : User_ClsCode()
'�ش���    : MDC_PS_Common
'��    ��    : ������ ������� ���ڵ� ��ȸ
'��    ��    : ����
'�� ȯ ��    : ClsCode
'Ư�̻���    : ����
'******************************************************************************
On Error GoTo User_ClsCode_Error

    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    sQry = "Select U_ClsCode From [OHEM] a Inner Join [OUSR] b On a.userId = b.USERID Where USER_CODE = '" & Trim(Sbo_Company.UserName) & "'"
    oRecordSet01.DoQuery sQry
    
    User_ClsCode = Trim(oRecordSet01.Fields(0).VALUE)
    Set oRecordSet01 = Nothing
    Exit Function

User_ClsCode_Error:
    Set oRecordSet01 = Nothing
    User_ClsCode = 0
    MDC_Com.MDC_GF_Message "User_ClsCode_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Public Function User_SuperUserYN() As String
'******************************************************************************
'Function ID : User_SuperUserYN()
'�ش���    : MDC_PS_Common
'��    ��    : ������ ������� SuperUser ����
'��    ��    : ����
'�� ȯ ��    : Y:��������, N:�Ϲ�����
'Ư�̻���    : ����
'******************************************************************************
On Error GoTo User_SuperUserYN_Error

    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    sQry = "           SELECT      T0.SUPERUSER"
    sQry = sQry & " FROM       OUSR AS T0"
    sQry = sQry & " WHERE      T0.User_Code = '" & Trim(Sbo_Company.UserName) & "'"
    
    Call oRecordSet01.DoQuery(sQry)
    
    User_SuperUserYN = Trim(oRecordSet01.Fields(0).VALUE)
    Set oRecordSet01 = Nothing
    Exit Function

User_SuperUserYN_Error:
    Set oRecordSet01 = Nothing
    User_SuperUserYN = 0
    MDC_Com.MDC_GF_Message "User_SuperUserYN_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Public Function User_MainJob() As String
'******************************************************************************
'Function ID : User_MainJob()
'�ش���    : MDC_PS_Common
'��    ��    : ������ ������� �ֿ���� ��ȸ
'��    ��    : ����
'�� ȯ ��    : �ֿ����(�λ縶����(OHEM)�� Remark)
'Ư�̻���    : ����
'******************************************************************************
On Error GoTo User_MainJob_Error

    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    sQry = "           SELECT       T0.Remark"
    sQry = sQry & " FROM        OHEM AS T0"
    sQry = sQry & "                 LEFT JOIN"
    sQry = sQry & "                 OUSR AS T1"
    sQry = sQry & "                     ON T0.UserID = T1.USERID"
    sQry = sQry & " WHERE       T1.User_Code = '" & Trim(Sbo_Company.UserName) & "'"
    
    Call oRecordSet01.DoQuery(sQry)
    
    User_MainJob = Trim(oRecordSet01.Fields(0).VALUE)
    Set oRecordSet01 = Nothing
    Exit Function

User_MainJob_Error:
    Set oRecordSet01 = Nothing
    User_MainJob = ""
    MDC_Com.MDC_GF_Message "User_MainJob_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Public Function Calculate_Weight(ByVal ItemCode$, ByVal Qty&, ByVal BPLID$) As Double
On Error GoTo Calculate_Weight_Error

    Dim ReturnValue  As Double
    Dim sQry         As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    sQry = "Select U_OBasUnit, U_UnitQ1, U_Spec1, U_Spec2, U_Spec3, U_UnWeight From [OITM] Where ItemCode = '" & ItemCode & "'"
    oRecordSet01.DoQuery sQry
    
    If Trim(oRecordSet01.Fields(0).VALUE) = "101" Then
        ReturnValue = Qty
    ElseIf Trim(oRecordSet01.Fields(0).VALUE) = "102" Then
        ReturnValue = Qty * Trim(oRecordSet01.Fields(1).VALUE)
    ElseIf Trim(oRecordSet01.Fields(0).VALUE) = "201" Then
        ReturnValue = (Trim(oRecordSet01.Fields(2).VALUE) - Trim(oRecordSet01.Fields(3).VALUE)) * Trim(oRecordSet01.Fields(3).VALUE) * 0.02808 * (Trim(oRecordSet01.Fields(4).VALUE) / 1000) * Qty
    ElseIf Trim(oRecordSet01.Fields(0).VALUE) = "202" Then
        ReturnValue = Qty * Trim(oRecordSet01.Fields(5).VALUE) / 1000
    ElseIf Trim(oRecordSet01.Fields(0).VALUE) = "203" Then
        ReturnValue = 0
    End If
    
    If BPLID = "3" Or BPLID = "5" Then
        Calculate_Weight = Round(ReturnValue, 2)
    Else
        Calculate_Weight = Round(ReturnValue, 0)
    End If
    
    Set oRecordSet01 = Nothing
    Exit Function

Calculate_Weight_Error:
    Set oRecordSet01 = Nothing
    Calculate_Weight = 0
    MDC_Com.MDC_GF_Message "Calculate_Weight_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Public Function Calculate_Qty(ByVal ItemCode$, ByVal Weight&) As Long
On Error GoTo Calculate_Qty_Error

    Dim ReturnValue  As Double
    Dim sQry         As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    sQry = "Select U_OBasUnit, U_UnitQ1, U_Spec1, U_Spec2, U_Spec3, U_UnWeight From [OITM] Where ItemCode = '" & ItemCode & "'"
    oRecordSet01.DoQuery sQry
    
    If Trim(oRecordSet01.Fields(0).VALUE) = "101" Then
        ReturnValue = Weight
    ElseIf Trim(oRecordSet01.Fields(0).VALUE) = "102" Then
        If Trim(oRecordSet01.Fields(1).VALUE) = "" Or Trim(oRecordSet01.Fields(1).VALUE) = 0 Then
            ReturnValue = 0
        Else
            ReturnValue = Weight / Trim(oRecordSet01.Fields(1).VALUE)
        End If
    ElseIf Trim(oRecordSet01.Fields(0).VALUE) = "201" Then
        If (Trim(oRecordSet01.Fields(2).VALUE) - Trim(oRecordSet01.Fields(3).VALUE)) * Trim(oRecordSet01.Fields(3).VALUE) * 0.02808 * (Trim(oRecordSet01.Fields(4).VALUE) / 1000) = "" Or _
        (Trim(oRecordSet01.Fields(2).VALUE) - Trim(oRecordSet01.Fields(3).VALUE)) * Trim(oRecordSet01.Fields(3).VALUE) * 0.02808 * (Trim(oRecordSet01.Fields(4).VALUE) / 1000) = 0 Then
            ReturnValue = 0
        Else
            ReturnValue = Weight / ((Trim(oRecordSet01.Fields(2).VALUE) - Trim(oRecordSet01.Fields(3).VALUE)) * Trim(oRecordSet01.Fields(3).VALUE) * 0.02808 * (Trim(oRecordSet01.Fields(4).VALUE) / 1000))
        End If
    ElseIf Trim(oRecordSet01.Fields(0).VALUE) = "202" Then
        If Trim(oRecordSet01.Fields(5).VALUE) = "" Or Trim(oRecordSet01.Fields(5).VALUE) = 0 Then
            ReturnValue = 0
        Else
            ReturnValue = Weight / Trim(oRecordSet01.Fields(5).VALUE) * 1000
        End If
    ElseIf Trim(oRecordSet01.Fields(0).VALUE) = "203" Then
        ReturnValue = 0
    End If
    
    Calculate_Qty = Round(ReturnValue, 0)
    Set oRecordSet01 = Nothing
    Exit Function

Calculate_Qty_Error:
    Set oRecordSet01 = Nothing
    Calculate_Qty = 0
    MDC_Com.MDC_GF_Message "Calculate_Qty_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Public Function RFC_Sender(ByVal BPLID As String, ByVal ItemCode As String, ByVal ItemName As String, ByVal Size As String, ByVal Qty As Double, ByVal Unit As String, ByVal RequestDate As String, ByVal DueDate As String, ByVal ItemType As String, ByVal RequestNo As String, ByVal i&, ByVal LastRow&) As String
On Error GoTo RFC_Sender_Error

    Dim ReturnValue As String
    Dim WERKS As String
       
    If i = 0 Then
        Set oSapConnection01 = CreateObject("SAP.Functions")
        oSapConnection01.Connection.User = "ifuser"
        oSapConnection01.Connection.Password = "pdauser"
'        oSapConnection01.Connection.client = "710"
        oSapConnection01.Connection.Client = "210"
'        oSapConnection01.Connection.ApplicationServer = "192.1.11.7"
        oSapConnection01.Connection.ApplicationServer = "192.1.1.217"
        oSapConnection01.Connection.language = "KO"
        oSapConnection01.Connection.SystemNumber = "00"
        If Not oSapConnection01.Connection.Logon(0, True) Then
            MDC_Com.MDC_GF_Message "�Ȱ�(R/3)������ �����Ҽ� �����ϴ�.", "E"
            GoTo RFC_Sender_Exit
        End If
    End If
    
    Dim oFunction01 As Object
    Set oFunction01 = oSapConnection01.Add("ZMM_INTF_GROUP")
    If BPLID = 1 Then
        WERKS = "9200"
    ElseIf BPLID = 2 Then
        WERKS = "9300"
    Else
        WERKS = "9200"
    End If
    
    oFunction01.Exports("I_WERKS") = WERKS '//�÷�Ʈ Ȧ���� â�� 9200, Ȧ���� �λ� 9300
    oFunction01.Exports("I_MATNR") = ItemCode '//�����ڵ� char(18)
    oFunction01.Exports("I_MAKTX") = ItemName '//���系�� char(40)
    oFunction01.Exports("I_WRKST") = Size '//����԰� char(48)
    oFunction01.Exports("I_MENGE") = Qty '//���ſ�û���� dec(13,3)
    oFunction01.Exports("I_MEINS") = Unit '//���� char(3)
    oFunction01.Exports("I_BADAT") = RequestDate '//���ſ�û�� char(8)
    oFunction01.Exports("I_LFDAT") = DueDate '//��ǰ�� char(8)
    oFunction01.Exports("I_MATKL") = ItemType '//����׷� char(9)
    oFunction01.Exports("I_ZBANFN") = RequestNo '//���ſ�û��ȣ

    If Not (oFunction01.Call) Then
        MDC_Com.MDC_GF_Message "�Ȱ�(R/3)���� �Լ�ȣ���� �����߻�", "E"
        GoTo RFC_Sender_Exit
    Else
        If (oFunction01.Imports("E_MESSAGE").VALUE = "") Then '//�����޽���
            ReturnValue = oFunction01.Imports("E_BANFN").VALUE & "/" & oFunction01.Imports("E_BNFPO").VALUE '//���ձ��ſ�û��ȣ '//���ձ��ſ�û ǰ���ȣ
        Else
            Call MDC_Com.MDC_GF_Message(oFunction01.Imports("E_MESSAGE").VALUE, "E")
            GoTo RFC_Sender_Exit
        End If
    End If
    
    If Not (oSapConnection01.Connection Is Nothing) Then
        If i = LastRow Then
            oSapConnection01.Connection.Logoff
            Set oSapConnection01 = Nothing
        End If
    End If
    
    RFC_Sender = ReturnValue
    Set oFunction01 = Nothing
    Exit Function
RFC_Sender_Exit:
    If Not (oSapConnection01.Connection Is Nothing) Then
        If i = LastRow Then
            oSapConnection01.Connection.Logoff
            Set oSapConnection01 = Nothing
        End If
    End If
    RFC_Sender = ""
    Set oFunction01 = Nothing
    Exit Function
RFC_Sender_Error:
    If Not (oSapConnection01.Connection Is Nothing) Then
        If i = LastRow Then
            oSapConnection01.Connection.Logoff
            Set oSapConnection01 = Nothing
        End If
    End If
    RFC_Sender = ""
    Set oFunction01 = Nothing
    Sbo_Application.SetStatusBarMessage "RFC_Sender_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Public Function Cal_KPI_Grade(ByVal prmBaseEntry As Integer, ByVal prmBaseLine As Integer, ByVal prmTableName As String, ByVal prmResult As String, ByVal prmMonth As String) As String
'******************************************************************************
'Function ID : Cal_KPI_Grade()
'�ش���    : MDC_PS_Common
'��    ��    : KPI �򰡵�� ���
'��    ��    : prmBaseEntry(KPI��ǥ������ȣ), prmBaseLine(KPI��ǥ�������ȣ), prmTableName(KPI��ǥ ���̺� ��), prmResult(����), prmMonth(������Ͽ�)
'�� ȯ ��    : KPI�򰡵��
'Ư�̻���    : ����
'******************************************************************************
On Error GoTo Cal_KPI_Grade_Error

    '1. �ش�KPI��ǥ ���̺��� ������ȣ�� ���ȣ�� �̿��Ͽ� A~E������ �� ��ȸ
    '2. ��ޱ���(�ִ�, �ּ�)�� ���� �б⹮�� �޶����� �ϹǷ� ��ޱ����� �ִ�����, �ּ����� �Բ� ��ȸ
    
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    sQry = "EXEC PS_Z_GetKPIGrade " & prmBaseEntry & "," & prmBaseLine & ",'" & prmTableName & "','" & prmResult & "', '" & prmMonth & "'"
    
    Call oRecordSet01.DoQuery(sQry)
    
    Cal_KPI_Grade = oRecordSet01.Fields("Grade").VALUE

    Set oRecordSet01 = Nothing
    Exit Function

Cal_KPI_Grade_Error:

    Cal_KPI_Grade = ""
    Set oRecordSet01 = Nothing
    Sbo_Application.SetStatusBarMessage "Cal_KPI_Grade_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    
End Function


Public Function Cal_KPI_Score(ByVal prmKPIGrade As String) As Double
'******************************************************************************
'Function ID : Cal_KPI_Score()
'�ش���    : MDC_PS_Common
'��    ��    : KPI ������ ���
'��    ��    : prmKPIGrade(KPI�򰡵��)
'�� ȯ ��    : KPI������
'Ư�̻���    : ����
'******************************************************************************
On Error GoTo Cal_KPI_Score_Error

    Dim sQry        As String
    Dim KPI_Score   As Double
    
    Dim loopCount01 As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    sQry = "        SELECT      T1.U_CodeNm AS [CodeName],"
    sQry = sQry & "             T1.U_Num1 AS [KPIScore]"
    sQry = sQry & " FROM        [@PS_HR200H] AS T0"
    sQry = sQry & "             INNER JOIN"
    sQry = sQry & "             [@PS_HR200L] AS T1"
    sQry = sQry & "                 ON T0.Code = T1.Code"
    sQry = sQry & " WHERE       T0.Name = '������'"
    
    Call oRecordSet01.DoQuery(sQry)
    
    For loopCount01 = 0 To oRecordSet01.RecordCount - 1
        
        If prmKPIGrade = oRecordSet01.Fields("CodeName").VALUE Then
        
            KPI_Score = oRecordSet01.Fields("KPIScore").VALUE
        
        End If
        
        oRecordSet01.MoveNext
    
    Next
    
    Cal_KPI_Score = KPI_Score
    
    Set oRecordSet01 = Nothing
    Exit Function
    
Cal_KPI_Score_Error:

    Set oRecordSet01 = Nothing
    Sbo_Application.SetStatusBarMessage "Cal_KPI_Score_Error " & Err.Number & " - " & Err.Description, bmt_Short, True

End Function

Public Function Cal_KPI_AchieveRate(ByVal prmBasEntry As Integer, ByVal prmBasLine As Integer, ByVal prmDocType As String, ByVal prmRslt As String) As Double
'******************************************************************************
'Function ID : Cal_KPI_AchieveRate()
'�ش���    : MDC_PS_Common
'��    ��    : KPI ��ô��(�޼���)
'��    ��    : prmBasEntry(��ǥ������ȣ), prmBasLine(��ǥ���ȣ), prmDocType(����Ÿ��), prmRslt(����)
'�� ȯ ��    : KPI������
'Ư�̻���    : ����
'******************************************************************************
On Error GoTo Cal_KPI_AchieveRate_Error

    Dim sQry         As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    sQry = "EXEC PS_Z_GetAchieveRate " & prmBasEntry & "," & prmBasLine & ",'" & prmDocType & "','" & prmRslt & "'" '��ô�� ��� ���ν���
    
    Call oRecordSet01.DoQuery(sQry)
    
    Cal_KPI_AchieveRate = oRecordSet01.Fields("AchieveRate").VALUE
    
    Set oRecordSet01 = Nothing
    Exit Function

Cal_KPI_AchieveRate_Error:

    Set oRecordSet01 = Nothing
    Sbo_Application.SetStatusBarMessage "Cal_KPI_AchieveRate_Error " & Err.Number & " - " & Err.Description, bmt_Short, True

End Function

Public Function Check_Finish_Status(ByVal prmBPLId As String, ByVal prmDocDate As String, ByVal prmFormTypeEx) As Boolean
'******************************************************************************
'Function ID : Check_Finish_Status()
'�ش���    : MDC_PS_Common
'��    ��    : �������� ��ȸ
'��    ��    : prmBPLID(�����), prmDocDate(�����), prmFormTypeEx(ȭ��Ÿ��(UID))
'�� ȯ ��    : �������¿� ���� ��� ���� ����
'Ư�̻���    : ����
'******************************************************************************
On Error GoTo Check_Finish_Status_Error
    
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim CheckFinishStatus As String
    
    sQry = "      EXEC PS_Z_CheckFinishStatus '"
    sQry = sQry & prmBPLId & "','"
    sQry = sQry & prmDocDate & "','"
    sQry = sQry & prmFormTypeEx & "'"

    Call oRecordSet01.DoQuery(sQry)
    
    CheckFinishStatus = oRecordSet01.Fields("ReturnValue").VALUE
    
    If CheckFinishStatus = "True" Then
        Check_Finish_Status = True
    Else
        Check_Finish_Status = False
    End If

    Set oRecordSet01 = Nothing

    Exit Function
    
Check_Finish_Status_Error:
    Set oRecordSet01 = Nothing
    Check_Finish_Status = False
    Call Sbo_Application.SetStatusBarMessage("Check_Finish_Status_Error " & Err.Number & " - " & Err.Description, bmt_Short, True)
End Function


