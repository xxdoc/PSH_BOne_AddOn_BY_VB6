Attribute VB_Name = "MDC_SetIni"
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
                                                                                          ByVal lpKeyName As Any, _
                                                                                          ByVal lpDefault As String, _
                                                                                          ByVal lpReturnedString As String, _
                                                                                          ByVal nSize As Long, _
                                                                                          ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
                                                                                              ByVal lpKeyName As Any, _
                                                                                              ByVal lpString As Any, _
                                                                                              ByVal lpFileName As String) As Long


Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

'�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
'���� �����ϴ� �Լ�
'�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
Private Declare Function SHFileOperation Lib "shell32.dll" _
Alias "SHFileOperationA" _
(lpFileOp As SHFILEOPSTRUCT) As Long

Private Const FO_COPY = &H2&

'�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
'���� ���ۿ� ���õ� ������ �����ϴ� ��������� ��������
'�ѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤѤ�
Private Type SHFILEOPSTRUCT
             hwnd As Long
             wfunc As Long
             pfrom As String
             pto As String
             fFlags As Long
             fAnyOperationsAborted As Long
             hNamemappings As Long
             lpszProgressTitle As String
End Type

Private Client_INI As String
Private Server_INI As String

Private ClientVer As String
Private ServerVer As String
Private PAY_ClientVer As String
Private PAY_ServerVer As String
Private ZPAY_DllName As String



Private dllUpdatePath As String

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
            ByVal hwnd As Long, ByVal lpOperation As String, _
            ByVal lpFile As String, ByVal lpParameters As String, _
            ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Const SW_HIDE = 0
Public Const SW_SHOWNORMAL = 1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_MINIMIZE = 6

Public Sub IniClient()
    '***************************************************************************
    ' INI�� ��ġ����
    '***************************************************************************
    Client_INI = App.Path & "\dllUpdate\ini\MDCINI.ini"
    SP_Path = GetSectionP("SetINI", "Path", Client_INI)
    SP_Screen = GetSectionP("SetINI", "ScreenName ", SP_Path)
    SP_Report = GetSectionP("SetINI", "ReportName ", SP_Path)  '3
End Sub

Private Sub File_Copy(sFileName, tFilename, mFilename)

    Dim utdShellOpStruct As SHFILEOPSTRUCT
    Dim lngReturnCode As Long
    Dim pValue As String
    Dim vpos As Integer
    Dim rtn As Integer
    
    '���� ���ۿ� ���� ������ ����
    With utdShellOpStruct
    
    '�ڵ鰪
    '.hWnd = frmLogin.hWnd
    
    '����
    .wfunc = FO_COPY
    
    '������ ����
    .pfrom = sFileName
    
    '������ ��ġ
    .pto = tFilename
    
    End With
    lngReturnCode = SHFileOperation(utdShellOpStruct)
    
    'rtn = Shell(mFilename, vbNormalFocus)

End Sub

Private Function GetSectionP(s As String, k As String, P As String) As String
    '***************************************************************************
    'Function  ID : GetSectionP
    '��        �� : INI���� ��������
    '��        �� : None
    '��   ȯ   �� : ��
    'Ư ��  �� �� : ���ǰ��� Ű�� �´� ���� ���������Ѵ�
    '***************************************************************************
    On Error GoTo Err
    Dim rtn_string As String * 255
    Dim rtn As Long
    rtn = GetPrivateProfileString(s, k, "", rtn_string, 255, P)
    If rtn = -1 Then
        Call Log("[" & s & "] " & k & " ���� ������ �о� �� ���� �����ϴ�.")
        GetSectionP = ""
    Else
        GetSectionP = Left(rtn_string, InStr(1, rtn_string, Chr(0)) - 1)
    End If
    Exit Function
Err:
    GetSectionP = ""
    MsgBox Err.Description
    Exit Function
End Function
Private Function SetSectionP(s As String, k As String, V As String, P As String) As Boolean
    '***************************************************************************
    'Function  ID : SetSectionP
    '��        �� : INI���� �����ֱ�
    '��        �� : None
    '��   ȯ   �� : None
    'Ư ��  �� �� : ���ǰ��� Ű�� �´� �� INI�� �����Ѵ�.
    '***************************************************************************
    On Error GoTo Err
    Dim rtn_string As String
    Dim rtn As Long
    Dim i As Integer
    rtn_string = V
    rtn = WritePrivateProfileString(s, k, rtn_string, P)
    If rtn = -1 Then
        Call Log("[" & s & "] " & k & " ���� ������ ���� �� ���� �����ϴ�.")
        SetSectionP = False
    Else
        SetSectionP = True
    End If
    Exit Function
Err:
    SetSectionP = False
    MsgBox Err.Description
    Exit Function
End Function




