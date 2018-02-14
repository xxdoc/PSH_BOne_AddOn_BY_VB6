VERSION 5.00
Begin VB.Form ZP_Form 
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   900
   ClientLeft      =   9840
   ClientTop       =   2550
   ClientWidth     =   1620
   LinkTopic       =   "Form1"
   ScaleHeight     =   900
   ScaleWidth      =   1620
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
End
Attribute VB_Name = "ZP_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type OPENFILENAME
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        Flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        IpTemplateName As String
 End Type
 
 Private Const OFN_READONLY = &H1
 Private Const OFN_OVERWRITEPROMPT = &H2
 Private Const OFN_HIDEREADONLY = &H4
 Private Const OFN_NOCHANGEDIR = &H8
 Private Const OFN_SHOWHELP = &H10
 Private Const OFN_ENABLEHOOK = &H20
 Private Const OFN_ENABLETEMPLATE = &H40
 Private Const OFN_ENABLETEMPLATEHANDLE = &H80
 Private Const OFN_NOVALIDATE = &H100
 
 Private Const OFN_ALLOWMULTISELECT = &H200
 Private Const OFN_EXTENSIONDIFFERENT = &H400
 Private Const OFN_PATHMUSTEXIST = &H800
 Private Const OFN_FILEMUSTEXIST = &H1000
 Private Const OFN_CREATEPROMPT = &H2000
 Private Const OFN_SHAREAWARE = &H400
 Private Const OFN_NOREADONLYRETURN = &H8000
 Private Const OFN_NOTESTFILECREATE = &H10000
 Private Const OFN_NONENETWORKBUTTON = &H20000
 Private Const OFN_NOLONGNAMES = &H40000
 Private Const OFN_EXPLORER = &H80000
 Private Const OFN_NODEREFERENCELINKS = &H100000
 Private Const OFN_LONGNAMES = &H20000
 
 Private Const OFN_SHAREFALLTHROUGH = 2
 Private Const OFN_SHARENOWARN = 1
 Private Const OFN_SHAREWARN = 0
 
Private Const conHwndTopmost = -1
Private Const conHwndNoTopmost = -2
Private Const conSwpNoActivate = &H10
Private Const conSwpShowWindow = &H40
 
'// 폴더선택
Private Const BIF_RETURNONLYFSDIRS = &H1

Private Type BROWSEINFO
  hOwner          As Long
  pidlRoot        As Long
  pszDisplayName  As String
  lpszTitle       As String
  ulFlags         As Long
  lpfn            As Long
  lParam          As Long
  iImage          As Long
End Type

 
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
'-------------------------------------------------------------------------------------------
' 폴더 찾아보기 창 관련 함수, 변수 및 상수
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" _
       (lpBrowseInfo As BROWSEINFO) As Long
       
' 드라이브 및 폴더명 가져오기 함수
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" _
       (ByVal pidl As Long, ByVal pszPath As String) As Long
       


Function OpenDialog(Form1 As Form, Filter As String, Title As String, InitDir As String)
    Dim ofn As OPENFILENAME
    Dim A As Long
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = Form1.hwnd
    ofn.hInstance = App.hInstance
    If Right$(Filter, 1) <> "|" Then Filter = Filter + "|"
    For A = 1 To Len(Filter)
        If Mid$(Filter, A, 1) = "|" Then Mid$(Filter, A, 1) = Chr$(0)
    Next
    
    ofn.lpstrFilter = Filter
    ofn.lpstrFile = Space$(254)
    ofn.nMaxFile = 255
    ofn.lpstrFileTitle = Space$(254)
    ofn.nMaxFileTitle = 255
    ofn.lpstrInitialDir = InitDir
    ofn.lpstrTitle = Title
    ofn.Flags = OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST
    A = GetOpenFileName(ofn)
    
    If (A) Then
        OpenDialog = Trim$(ofn.lpstrFile)
    
    Else
        OpenDialog = ""
    End If

End Function
 
' 폴더찾아보기 창 띄우기
Function vbGetBrowseDirectory(Form1 As Form) As String
    Dim bi As BROWSEINFO
    Dim r As Long
    Dim pidl As Long         '폴더찾아보기에서 리턴값
    Dim tmpPath As String
    
    '---------------------------------------------------
    ' 폴더찾아보기 창 띄움
    bi.hOwner = Form1.hwnd
    bi.pidlRoot = 0&       'vbNull
    bi.lpszTitle = "드라이브 및 폴더를 선택하세요.."
    bi.ulFlags = BIF_RETURNONLYFSDIRS
    
    pidl = SHBrowseForFolder(bi)   '폴더찾아보기 창 띄움 0:취소버튼클릭

    '---------------------------------------------------
    '선택된 드라이브 및 폴더명 가져오기
    tmpPath = Space(512)
    r = SHGetPathFromIDList(ByVal pidl, ByVal tmpPath)  '1:확인버튼클릭,  0:취소버튼클릭
    If r > 0 Then
          vbGetBrowseDirectory = Left(tmpPath, InStr(tmpPath, Chr(0)) - 1)
    Else
          vbGetBrowseDirectory = ""
    End If
End Function
Private Sub Form_Load()
    SetWindowPos hwnd, conHwndTopmost, 0, 0, 0, 0, _
        conSwpNoActivate Or conSwpShowWindow
End Sub

