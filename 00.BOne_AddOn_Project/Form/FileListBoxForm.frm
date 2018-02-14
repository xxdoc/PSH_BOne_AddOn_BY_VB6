VERSION 5.00
Begin VB.Form FileListBoxForm 
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   2910
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2295
   LinkTopic       =   "Form1"
   ScaleHeight     =   2910
   ScaleWidth      =   2295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
End
Attribute VB_Name = "FileListBoxForm"
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
 
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long

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
Private Sub Form_Load()
    SetWindowPos hwnd, conHwndTopmost, 0, 0, 0, 0, _
        conSwpNoActivate Or conSwpShowWindow
End Sub





