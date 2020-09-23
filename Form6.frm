VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form6 
   BackColor       =   &H80000000&
   BorderStyle     =   0  'None
   Caption         =   "Fonix - Changing Status"
   ClientHeight    =   645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3495
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   645
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   -405
      Top             =   0
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   255
      Left            =   75
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Label Label1 
      Caption         =   "Changing status. Please wait ..."
      ForeColor       =   &H80000001&
      Height          =   255
      Left            =   675
      TabIndex        =   1
      Top             =   405
      Width           =   2535
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a As Integer
Dim b As Integer

Private Sub Form_Load()
a = 0
b = 0
End Sub

Private Sub Timer1_Timer()
If Form1.CurrStatus = True Then
    Form1.CurrStatus = False
    b = 1
    pb.Value = 100
    a = 0
    Unload Me
End If


If a > 90 And b <> 1 Then
    a = 90
    b = 0
End If
pb.Value = a
a = a + 5
End Sub
