Attribute VB_Name = "MessageAPIs"


'//  SAP MANAGE UI API 6.7 SDK Sample
'//****************************************************************************
'//
'//  File:      MessageAPIs.bas
'//
'//  Copyright (c) SAP MANAGE
'//
'// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF
'// ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
'// THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
'// PARTICULAR PURPOSE.
'//
'//****************************************************************************

Option Explicit
'//****************************************************************
'// API Declarations
'// Enable us to process a message loop in Sub Main()
'//
'// A developer should copy this module 'as is' and create
'// an object of your class in Sub Main()
'//****************************************************************

'// Part of the MSG structure - receives the location of the mouse
Public Type POINTAPI
    x As Long
    y As Long
End Type

'// The message structure
Public Type Msg
    hWnd As Long
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
      ByVal hWnd As Long, _
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


Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hWnd As Long, ByVal IpOperation As String, _
 ByVal IpFile As String, ByVal IpParameters As String, _
 ByVal IpDirectory As String, ByVal nShowCmd As Long) As Long


