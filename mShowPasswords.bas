Attribute VB_Name = "mShowPasswords"
'/*
'**   This piece of code shows text in all open windows that have masked
'**   password boxes on them.
'**
'**   For instance, your dialup network login box displays ********
'**   when you call the function in this module, it will display
'**   the actual password; for instance 'pass1243'
'**
'**
'**   written by SubReality.
'**   Downloaded from http://www.subreality.net
'*/


Option Explicit
'/* API Calls...
'*/
Declare Function EnumChildWindows Lib "user32" ( _
        ByVal hWndParent As Long, _
        ByVal lpEnumFunc As Long, _
        ByVal lParam As Long) As Long
Declare Function EnumWindows Lib "user32" ( _
        ByVal lpEnumFunc As Long, _
        ByVal lParam As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
        ByVal hWnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Integer, _
        ByVal lParam As Long) As Long
Public Declare Function ShowWindow Lib "user32" ( _
        ByVal hWnd As Long, _
        ByVal nCmdShow As Long) As Long

Private Const EM_GETPASSWORDCHAR = &HD2
Private Const EM_SETPASSWORDCHAR = &HCC
Private Const EM_SETMODIFY = &HB9
Private Const SW_HIDE = 0
Private Const SW_SHOW = 5

'/* Callback for parent windows...
'*/
Private Function EnumWindowsProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
  EnumChildWindows hWnd, AddressOf EnumWindowsProc2, 1
  EnumWindowsProc = True
End Function
'/* Callback for child windows...
'*/
Private Function EnumWindowsProc2(ByVal hWnd As Long, ByVal lParam As Long) As Long
  If SendMessage(hWnd, EM_GETPASSWORDCHAR, 0, 1) Then
   UpdateWindow hWnd
  End If
  EnumWindowsProc2 = True
End Function

'/* Show the password and refresh textbox...
'*/
Private Sub UpdateWindow(hWnd As Long)
  SendMessage hWnd, EM_SETPASSWORDCHAR, 0, 1
  SendMessage hWnd, EM_SETMODIFY, True, 1
  ShowWindow hWnd, SW_HIDE
  ShowWindow hWnd, SW_SHOW
End Sub

'/* the only Public function...
'*/
Public Function UnmaskPasswords()
  EnumWindows AddressOf EnumWindowsProc, 1
End Function
