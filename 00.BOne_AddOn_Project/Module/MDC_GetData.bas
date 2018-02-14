Attribute VB_Name = "MDC_GetData"
'*******************************************************************************
' ȭ��  ID : MDC_GetData
' ȭ �� �� :
' ��    �� : ���(����ó���ÿ� ������ ��ȯ�ϴ� �Լ��� ����)
' Table �� : None
' �Է�  �� :
' ���  �� :
' �� �� �� : �迵ȣ,�Թ̰�
' �� �� �� : 2005. 08. 22~~~~~~~~~~~~~~~~~~
'//  Copyright  (c) Morning Data
'-------------------------------------------------------------------------------
' �� �� ��    |    �� �� ��    |                   �� �� �� ��
'-------------------------------------------------------------------------------
Option Explicit

Public Sub SpaceLineDelete(TableId$, KeyColumn$)
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    '���̺� ����ִ� ������ �����Ѵ�
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry$
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    sQry = "Delete " & TableId
    sQry = sQry + " WHERE ISNULL(" & KeyColumn & ",'')=''"
    oRecordSet.DoQuery sQry
    
    Set oRecordSet = Nothing
End Sub

Public Sub SetMessage_Err(Msg As String)  '�����޼��� ���
    Call Sbo_Application.SetStatusBarMessage(Msg, bmt_Short, True)
End Sub

Public Sub SetMessage_Basic(Msg As String)   '�����޼��� ���
    Call Sbo_Application.SetStatusBarMessage(Msg, bmt_Short, False)
End Sub

Public Function Get_ItemName(ItemCode$) As String
'// ǰ���� ��ȯ�մϴ�
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
    Sbo_Application.SetStatusBarMessage "ǰ���� �����ü� �����ϴ�." & Space(10) & Err.Description, bmt_Short, True
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
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
    '��ȯ�÷�,���� �÷�,���̺�,���ǰ�,�ص���
    '�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
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
'/ ��������� DocEntry
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
'/ ���̺��̳� ���������� �ִ��� üũ�Ѵ�.
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
'/ �ش� Ŭ������ ���̺��� üũ
'-----------------------------------------
Public Function ObjectTableChk(ByVal Tablename$) As Boolean
On Error GoTo ObjectTableChk_Error:
    
    ObjectTableChk = True

    If ObjectExitsChk(Tablename) = False Then
        If Z_Language = 28 Then
            Sbo_Application.SetStatusBarMessage " ���̺�  [" & Tablename & "] �� �ʿ��մϴ�  ! ", bmt_Short, False
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
'/ �ش� Ŭ������ ������Ʈ��� üũ
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
            Sbo_Application.SetStatusBarMessage Tablename & " �� UDO ����� �ʿ��մϴ�  ! ", bmt_Short, False
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
'/ ���ڸ� �ѱ۷� ǥ��
'-----------------------------------------
'Public Function NumberToHanGul(num) As String
'On Error GoTo NumberToHanGul_Error
'     Dim i As Integer, size As Integer, dsize As Integer
'     Dim sw As Boolean
'     Dim tempstring As String
'     Dim temp1 As String, temp2 As String
'     Dim DigitArry1, DigitArry2, NumberArry
'
'     DigitArry1 = Array("", "��", "��", "õ")
'     DigitArry2 = Array("", "��", "��", "��")
'     NumberArry = Array("", "��", "��", "��", "��", "��", "��", "ĥ", "��", "��")
'     If Application.IsNumber(num) = False Then
'        'NumberToHanGul = "���ڰ� �ƴմϴ�."
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
'     '���⼭ �ʿ��� ���ξ ���̾ ���ϼ� �ֽ��ϴ�
'     '  NumberToHanGul = "�ϱ�" & tempstring & "��"
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
'    Dim �����迭1, �����迭2, ���ڹ迭
'
' �����迭1 = Array("", "��", "��", "õ")
' �����迭2 = Array("", "��", "��", "��")
' ���ڹ迭 = Array("", "��", "��", "��", "��", "��", "��", "ĥ", "��", "��")
'
' If Application.IsNumber(num) = False Then
'    NumberToHanGul = "���ڰ� �ƴմϴ�."
'    Exit Function
' End If
' If num = 0 Then Exit Function
'
' size = Len(num)
' sw = True
'
' For i = 1 To size
'   dsize = size - i
'   temp1 = ���ڹ迭(CInt(Mid(num, i, 1)))
'   If temp1 <> "" Then
'      If dsize Mod 4 <> 0 Then
'         temp2 = �����迭1(dsize Mod 4)
'         sw = True
'      Else
'         temp2 = �����迭2(dsize \ 4)
'         sw = False
'      End If
'   Else
'      If dsize Mod 4 <> 0 Then
'         temp2 = ""
'      Else
'         If sw = True Then
'            temp2 = �����迭2(dsize \ 4)
'            sw = False
'         End If
'      End If
'   End If
'   tempstring = tempstring & temp1 & temp2
' Next i
' '���⼭ �ʿ��� ���ξ ���̾ ���ϼ� �ֽ��ϴ�
' '  NumberToHanGul = "�ϱ�" & tempstring & "��"
' NumberToHanGul = tempstring
'
'End Function

