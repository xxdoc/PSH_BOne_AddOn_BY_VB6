Attribute VB_Name = "MessageAPIs"
Option Explicit

'// Part of the MSG structure - receives the location of the mouse
Public Type POINTAPI
    x As Long
    y As Long
End Type

'// The message structure
Public Type Msg
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    Time As Long
    pt As POINTAPI
End Type

'// Retrieves messages sent to the calling thread's message queue
Public Declare Function GetMessage Lib "user32" _
    Alias "GetMessageA" _
     (lpMsg As Msg, _
      ByVal hwnd As Long, _
      ByVal wMsgFilterMin As Long, _
      ByVal wMsgFilterMax As Long) As Long
      
'// Translates virtual-key messages into character messages
Public Declare Function TranslateMessage Lib "user32" _
    (lpMsg As Msg) As Long

'// Forwards the message on to the window represented by the
'// hWnd member of the Msg structure
Public Declare Function DispatchMessage Lib "user32" _
    Alias "DispatchMessageA" _
     (lpMsg As Msg) As Long

Public Msg As Msg


Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long

Public Const MAX_PATH = 260             '//���ϸ��� �ִ� ����

Type FILETIME       '//���Ϲ��� ���Ͽ� ���� �Ϻ� �������� ����ü
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Type WIN32_FIND_DATA    '//�˻��� ���� �Ǵ� �������丮�� ������ ���� ����ü
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type


'���� �����ϴ� �Լ�
Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

'���� ���ۿ� ���õ� ������ �����ϴ� ��������� ��������
Type SHFILEOPSTRUCT
   hwnd                    As Long
   wfunc                   As Long
   pfrom                   As String
   pto                     As String
   fFlags                  As Long
   fAnyOperationsAborted   As Long
   hNamemappings           As Long
   lpszProgressTitle       As String
End Type


Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)




