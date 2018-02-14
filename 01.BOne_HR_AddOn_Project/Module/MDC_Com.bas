Attribute VB_Name = "MDC_Com"
Option Explicit
Public Type MonthDate
    StDate As String '* 10 'YYYY-MM-DD
    EdDate As String '* 10  'YYYY-MM-DD
End Type


Public Type PerSonNo_Info
    check As Boolean
    BirthDay As String
    Sex As Integer     '�� : 0, �� : 1
    ManAge As Integer
    Age As Integer
End Type

Public Function GetSpStr(Str As String) As String
    'Function ID : GetSpStr
    '��    ��    : ��ĭ �ٿ��� ���� �տ����� �����´�
    '��    ��    : Str
    '�� ȯ ��    : None
    'Ư�̻���    : ex) Str = 0 �����, ���� = �����
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
    '��    ��    : ��ĭ �ٿ��� ���� �ڿ����� �����´�
    '��    ��    : Str
    '�� ȯ ��    : None
    'Ư�̻���    : ex) Str = 0 �����, ���� = 0
    Dim Buf As String

    Buf = GetSpStr(Str)

    If Len(Str) > Len(Buf) + 1 Then
        GetSpStr2 = Mid(Str, Len(Buf) + 2, Len(Str) - Len(Buf) - 1)
    End If
End Function

Public Function uISDATE(Dt As Variant, Conv As Variant) As Variant
    'Function ID : uISDATE
    '��    ��    : ��¥���� ������ ����
    '��    ��    : dt,Conv
    '�� ȯ ��    : None
    'Ư�̻���    : None
    If IsDate(Dt) Then
        uISDATE = Dt
    Else
        uISDATE = Conv
    End If
End Function

Public Function uISNULL(Str As Variant, Conv As Variant) As Variant
    'Function ID : uISNULL
    '��    ��    : �������� ������ ����
    '��    ��    : Str,Conv
    '�� ȯ ��    : None
    'Ư�̻���    : None
    uISNULL = IIf(IsNull(Str), Conv, Str)
End Function

Public Function uISNUMERIC(num As Variant, Conv As Variant, NumType As String) As Variant
    'Function ID : uISNUMERIC
    '��    ��    : ���������� ��ȯ
    '��    ��    : Str,Conv
    '�� ȯ ��    : None
    'Ư�̻���    : None
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
Function GetLeftNumZero(num As String, �ڸ��� As Long) As String
    'Function ID : GetLeftNumZero
    '��    ��    : ���ڸ� ���������ϸ鼭 ���ʿ� '0'�� ä��� �Լ�
    '��    ��    : Num,�ڸ���
    '�� ȯ ��    : None
    'Ư�̻���    :
    ' ��ex) GetLeftNumZero("123456",10) --> ��ȯ�� = "0000123456"
    Dim Ln As Long

    Ln = Len(num)
    If Ln <= �ڸ��� Then
        GetLeftNumZero = String(�ڸ��� - Ln, "0") + num
    Else
        GetLeftNumZero = String(�ڸ���, "0")
    End If
End Function

'-----------------------------------------------------------------------------------------
'   �׺���̼� ��Ʈ�� ���� ���̱�/���߱� �Լ�
'   -> �̸�����, ���, �����, ã��, �߰�, ����, ����, ��ó��, �ǳ�, ���
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
        MDC_eForm.EnableMenu "519", MDC_bPreview         '// �μ�[�̸�����]
        MDC_eForm.EnableMenu "520", MDC_bPrint           '// �μ�[���]
        MDC_eForm.EnableMenu "1293", MDC_bDeleteRow      '// �����
        MDC_eForm.EnableMenu "1281", MDC_bFind           '// ����ã��
        MDC_eForm.EnableMenu "1282", MDC_bAdd            '// �����߰�
        MDC_eForm.EnableMenu "1283", MDC_bRemove         '// ��������(������ ������ ���)
        MDC_eForm.EnableMenu "1286", MDC_bClose          '// �����ݱ�
        MDC_eForm.EnableMenu "1288", MDC_bNextRecord     '// ����
        MDC_eForm.EnableMenu "1289", MDC_bPreviousRecord '// ����
        MDC_eForm.EnableMenu "1290", MDC_bFirstRecord    '// ��ó��
        MDC_eForm.EnableMenu "1291", MDC_bLastRecord     '// �ǳ�
        MDC_eForm.EnableMenu "1284", MDC_bCancel         '// �������
        MDC_eForm.EnableMenu "1292", MDC_bRowAdd         '// ���߰�
        MDC_eForm.EnableMenu "1287", MDC_bDuplicate      '// ��������
        MDC_eForm.EnableMenu "1299", MDC_bRowClose       '// ��ݱ�
        MDC_eForm.EnableMenu "1285", MDC_bRestore
    '//End If
End Sub

'// Matrix Combo Box Setting
Public Sub MDC_GP_MatrixSetMatComboList(MDC_fCombo As SAPbouiCOM.Column, _
                                        MDC_fSQL As String, _
                                        Optional AndLine$, _
                                        Optional AddSpace$)
    'Function ID : GetListIndex
    '��    ��    :
    '��    ��    : Lst
    '�� ȯ ��    : None
    'Ư�̻���    : �޺��ڽ��� ���� �� ������ �ý��� �ڵ忡�� ������ �����Ѵ�
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
'//     �޼��� ��� ���
'--------------------------------------------------------------------------------------
Public Function MDC_GF_Message(MDC_pMsg As String, MDC_pType As String) As Long    '//����޼���
    
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
' �ֹε�Ϲ�ȣ/�ܱ��ε�Ϲ�ȣ ����üũ
' Argument  : sID : �����Ϸ��� ��ȣ
' Remark    : �ܱ��� �ֹι�ȣ : 000000-1234567(���� 6�ڸ��� �������,���ڸ��� 1:��������, 23:��ϱ����ȣ, 45:�Ϸù�ȣ, 6:����ڱ���, 7:������ȣ)
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
    
    '// ���������ڵ�(1,2,3,4:������, 5,6,7,8:�ܱ���)
    If Mid(sID, 7, 1) < "1" Or Mid(sID, 7, 1) > "8" Then Exit Function
    '// �����ڵ�
    Select Case Mid(sID, 7, 1)
    Case "5", "6", "7", "8"   '// �ܱ���
        If (Val(Mid(sID, 8, 2)) Mod 2) <> 0 Then   '// ��ϱ����ȣ����
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
    Case "5", "6", "7", "8"   '// �ܱ���
        Rmn = Rmn + 2
        If Rmn >= 10 Then Rmn = Rmn - 10
    End Select
    
    GovIDCheck = IIf(Rmn = chk, True, False)
End Function

Function Age_Chk(Str As String, StdDate$) As PerSonNo_Info
    '***************************************************************************
    'Function ID : �ֹι�ȣüũ ���̱���.
    '��    ��    :
    '��    ��    : ��ġ
    '�� ȯ ��    : None
    'Ư�̻���    : None
    '�߰�        : �Թ̰�
    '***************************************************************************
    Dim chk As PerSonNo_Info
    Dim PerNo As String
    Dim Bir_YY As String
    
    PerNo = Replace(Str, "-", "")  '/ �ֹε�Ϲ�ȣ
    With chk
        .check = True
        Select Case Mid$(PerNo, 7, 1)
        Case "1":        Bir_YY = "19": .Sex = 0
        Case "2":        Bir_YY = "19": .Sex = 1   '/ ����
        Case "5":        Bir_YY = "19": .Sex = 0
        Case "6":        Bir_YY = "19": .Sex = 1
        Case "3":        Bir_YY = "20": .Sex = 0
        Case "4":        Bir_YY = "20": .Sex = 1
        Case "7":        Bir_YY = "20": .Sex = 0
        Case "8":        Bir_YY = "20": .Sex = 1
        Case "9":        Bir_YY = "18": .Sex = 0
        Case "0":        Bir_YY = "18": .Sex = 1
        End Select
        .BirthDay = Format$(Trim$(Bir_YY) & Left$(PerNo, 6), "0000-00-00")   '/ ����
        If Len(PerNo) <> 13 Then .check = False
        '����üũ
        If IsDate(.BirthDay) = False Then    '/ ��¥ üũ
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
