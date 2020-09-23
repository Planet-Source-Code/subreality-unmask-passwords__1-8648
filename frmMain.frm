VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "SubReality.net - Password Unmasker example"
   ClientHeight    =   945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4920
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   945
   ScaleWidth      =   4920
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   885
      Left            =   2490
      TabIndex        =   2
      Top             =   0
      Width           =   2385
      Begin VB.CommandButton cmdWebsite 
         Caption         =   "Visit Web Site"
         Height          =   525
         Left            =   120
         TabIndex        =   3
         Top             =   210
         Width           =   2115
      End
   End
   Begin VB.Frame Frame1 
      Height          =   885
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   2385
      Begin VB.CommandButton cmdUnmask 
         Caption         =   "Unmask Passwords"
         Height          =   525
         Left            =   120
         TabIndex        =   1
         Top             =   210
         Width           =   2115
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'/* API for executing a file or location,
'** used for visiting www.subreality.net...
'*/
Private Declare Function ShellExecute Lib "shell32.dll" _
        Alias "ShellExecuteA" ( _
        ByVal hWnd As Long, _
        ByVal lpOperation As String, _
        ByVal lpFile As String, _
        ByVal lpParameters As String, _
        ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As Long

'/* This is all you have to do...
'*/
Private Sub cmdUnmask_Click()
  UnmaskPasswords
End Sub

'/* Check out http://www.subreality.net for more excellent source code...
'*/
Private Sub cmdWebsite_Click()
  Call ShellExecute(0&, vbNullString, "http://www.subreality.net", vbNullString, vbNullString, vbNormalFocus)
End Sub
