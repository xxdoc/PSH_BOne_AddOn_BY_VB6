VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPDFium"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'This Wrapper depends on the commercially usable (BSD-licensed) PDFium-library,
'... in this little viewer-demo I'm using the pre-compiled Windows-32bit-binary from:
'https://github.com/pvginkel/PdfiumViewer/tree/master/Libraries
Private Declare Function LoadLibraryW Lib "kernel32" (ByVal lpLibFileName As Long) As Long

Public Enum PDFPageRotation
  Rot_0
  Rot_90
  Rot_180
  Rot_270
End Enum

Private Declare Sub FPDF_InitLibrary Lib "PDFium" Alias "_FPDF_InitLibrary@4" (Optional ByVal hInstance As Long)
Private Declare Sub FPDF_DestroyLibary Lib "PDFium" Alias "_FPDF_DestroyLibrary@0" ()
Private Declare Function FPDF_LoadMemDocument Lib "PDFium" Alias "_FPDF_LoadMemDocument@12" (ByVal pData As Long, ByVal DataLen As Long, ByVal Password As String) As Long
Private Declare Sub FPDF_CloseDocument Lib "PDFium" Alias "_FPDF_CloseDocument@4" (ByVal hDoc As Long)
Private Declare Function FPDF_GetPageCount Lib "PDFium" Alias "_FPDF_GetPageCount@4" (ByVal hDoc As Long) As Long
Private Declare Function FPDF_LoadPage Lib "PDFium" Alias "_FPDF_LoadPage@8" (ByVal hDoc As Long, ByVal PageIdx As Long) As Long
Private Declare Sub FPDF_ClosePage Lib "PDFium" Alias "_FPDF_ClosePage@4" (ByVal hPage As Long)
Private Declare Sub FPDF_RenderPage Lib "PDFium" Alias "_FPDF_RenderPage@32" (ByVal hDC&, ByVal hPage&, ByVal x&, ByVal y&, ByVal dx&, ByVal dy&, ByVal Rotation As PDFPageRotation, ByVal Flags&)
Private Declare Function FPDFBitmap_Create Lib "PDFium" Alias "_FPDFBitmap_Create@12" (ByVal dx As Long, ByVal dy As Long, ByVal Alpha As Long) As Long
Private Declare Sub FPDF_RenderPageBitmap Lib "PDFium" Alias "_FPDF_RenderPageBitmap@32" (ByVal hBM&, ByVal hPage&, ByVal x&, ByVal y&, ByVal dx&, ByVal dy&, ByVal Rotation As PDFPageRotation, ByVal Flags&)
Private Declare Function FPDFBitmap_GetBuffer Lib "PDFium" Alias "_FPDFBitmap_GetBuffer@4" (ByVal hBM As Long) As Long
Private Declare Sub FPDFBitmap_Destroy Lib "PDFium" Alias "_FPDFBitmap_Destroy@4" (ByVal hBM As Long)
Private Declare Function FPDF_GetPageWidth Lib "PDFium" Alias "_FPDF_GetPageWidth@4" (ByVal hPage As Long) As Double
Private Declare Function FPDF_GetPageHeight Lib "PDFium" Alias "_FPDF_GetPageHeight@4" (ByVal hPage As Long) As Double

Private hDoc As Long, Content() As Byte 'as long as a document is open, this buffer should not be touched or changed
Private Pages() As Long, mPageCount As Long

Private Sub Class_Initialize()
  LoadLibraryW StrPtr(App.Path & "\PDFium\pdfium.dll")
  FPDF_InitLibrary
End Sub

Public Property Get PDFByteContent() As Byte()
  PDFByteContent = Content
End Property

Public Property Let PDFByteContent(NewContent() As Byte)
Dim i As Long, DataLen As Long
  CloseDocument
  Content = NewContent
  DataLen = UBound(Content) - LBound(Content) + 1
  hDoc = FPDF_LoadMemDocument(VarPtr(Content(LBound(Content))), DataLen, "")
  If hDoc = 0 Then Err.Raise vbObjectError, , "couldn't open PDF from content"
  
  mPageCount = FPDF_GetPageCount(hDoc)
  If mPageCount = 0 Then CloseDocument: Err.Raise vbObjectError, , "couldn't open any pages"
  ReDim Pages(0 To mPageCount - 1)
  For i = 0 To mPageCount - 1
    Pages(i) = FPDF_LoadPage(hDoc, i)
    If Pages(i) = 0 Then CloseDocument: Err.Raise vbObjectError, , "couldn't open page at Idx: " & i
  Next i
End Property

Public Property Get PageCount() As Long
  PageCount = mPageCount
End Property

Public Property Get PageWidthPoints(ByVal PageIdxZeroBased&) As Double
  PageWidthPoints = FPDF_GetPageWidth(Pages(PageIdxZeroBased))
End Property
Public Property Get PageHeightPoints(ByVal PageIdxZeroBased&) As Double
  PageHeightPoints = FPDF_GetPageHeight(Pages(PageIdxZeroBased))
End Property

Public Sub RenderPageToDC(ByVal hDC&, ByVal PageIdxZeroBased&, ByVal x&, ByVal y&, ByVal dx&, ByVal dy&, Optional ByVal Rotation As PDFPageRotation, Optional ByVal ClearType As Boolean = True, Optional ByVal ShowAnnotations As Boolean)
Dim Flags As Long
    Flags = IIf(ClearType, 2, 0) Or IIf(ShowAnnotations, 1, 0) 'FPDF_NO_GDIPLUS = 4 ... (could be used to speed things up on windows at the cost of quality)
    FPDF_RenderPage hDC, Pages(PageIdxZeroBased), x, y, dx, dy, Rotation, Flags
End Sub

Public Function RenderPageToSrf(ByVal PageIdxZeroBased&, ByVal x&, ByVal y&, ByVal dx&, ByVal dy&, Optional ByVal Rotation As PDFPageRotation, Optional ByVal ClearType As Boolean = True, Optional ByVal ShowAnnotations As Boolean) As cCairoSurface
Dim Flags As Long, hBM As Long, pData As Long
    If dx <= 0 Then dx = 1
    If dy <= 0 Then dy = 1
    Flags = IIf(ClearType, 2, 0) Or IIf(ShowAnnotations, 1, 0) 'FPDF_NO_GDIPLUS = 4 ... (could be used to speed things up on windows at the cost of quality)
    hBM = FPDFBitmap_Create(dx, dy, 1)
    If hBM = 0 Then Err.Raise vbObjectError, , "couldn't create FPDF-Alpha-Bitmap"
On Error GoTo 1:
    Set RenderPageToSrf = Cairo.CreateSurface(dx, dy)
    FPDF_RenderPageBitmap hBM, Pages(PageIdxZeroBased), x, y, dx, dy, Rotation, Flags
    pData = FPDFBitmap_GetBuffer(hBM)
    If pData Then New_c.MemCopy RenderPageToSrf.DataPtr, pData, dx * dy * 4
1: FPDFBitmap_Destroy hBM
End Function
 
Public Sub CloseDocument()
Dim i As Long
  If hDoc = 0 Then Exit Sub

  For i = 0 To mPageCount - 1: FPDF_ClosePage Pages(i): Next
  mPageCount = 0
  
  FPDF_CloseDocument hDoc
  hDoc = 0
End Sub

