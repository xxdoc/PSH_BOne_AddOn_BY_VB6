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

Public Const MAX_PATH = 260             '//파일명의 최대 길이

Type FILETIME       '//리턴받을 파일에 관한 일부 세부정보 구조체
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Type WIN32_FIND_DATA    '//검색된 파일 또는 하위디렉토리의 정보를 받을 구조체
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


'파일 조작하는 함수
Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

'파일 조작에 관련된 정보를 정의하는 사용자정의 데이터형
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




