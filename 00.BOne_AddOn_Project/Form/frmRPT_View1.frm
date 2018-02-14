VERSION 5.00
Object = "{3C62B3DD-12BE-4941-A787-EA25415DCD27}#10.0#0"; "crviewer.dll"
Begin VB.Form frmRPT_View1 
   Caption         =   "CrystalReport Viewer10"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12615
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   12615
   StartUpPosition =   3  'Windows 기본값
   WindowState     =   2  '최대화
   Begin CrystalActiveXReportViewerLib10Ctl.CrystalActiveXReportViewer CRViewer1 
      Height          =   7935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12615
      lastProp        =   600
      _cx             =   22251
      _cy             =   13996
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   -1  'True
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
   End
End
Attribute VB_Name = "frmRPT_View1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub
