VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "NT/2000 Trap Keys ©2000 Jeremy Stein"
   ClientHeight    =   1680
   ClientLeft      =   1965
   ClientTop       =   1545
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   ScaleHeight     =   1680
   ScaleWidth      =   6180
   Begin VB.CommandButton Command2 
      Caption         =   "Enable"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4440
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmMain.frx":0000
      Top             =   120
      Width           =   5895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Disable"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   1575
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Jeremy Stein

'After hunting around I found a good example on DevX that
'I was able to build from and make easy to understand.

'Have fun and vote please

Dim hHook As Long


Private Sub Command1_Click()
  
  hHook = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf KeyboardHookProc, App.hInstance, 0)
  Command1.Enabled = False
  Command2.Enabled = True
End Sub

Private Sub Command2_Click()
  
  UnhookWindowsHookEx hHook
  hHook = 0
  Command2.Enabled = False
  Command1.Enabled = True
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  If hHook <> 0 Then UnhookWindowsHookEx hHook
  
End Sub
