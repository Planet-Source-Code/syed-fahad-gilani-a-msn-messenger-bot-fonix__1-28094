VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3360
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   4065
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   4065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000004&
      Height          =   3090
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3825
      Begin VB.Timer Timer3 
         Interval        =   500
         Left            =   720
         Top             =   2160
      End
      Begin MSComctlLib.ProgressBar pb 
         Height          =   255
         Left            =   885
         TabIndex        =   4
         Top             =   2760
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Max             =   240
         Scrolling       =   1
      End
      Begin VB.Timer Timer2 
         Interval        =   1
         Left            =   3360
         Top             =   240
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   120
         Top             =   1560
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Loading ..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   2520
         Width           =   3255
      End
      Begin VB.Label Label2 
         Caption         =   "MSN Messenger                      Fonix Extensions"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   495
         Left            =   1080
         TabIndex        =   2
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "By: Syed Fahad Gilani"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   1
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Image Image1 
         Height          =   900
         Left            =   120
         Picture         =   "frmSplash.frx":000C
         Top             =   240
         Width           =   3540
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Long
Dim j As Integer

Private Sub Form_Activate()
i = 0
j = 0
Load Form1
End Sub

Private Sub Form_Unload(Cancel As Integer)
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
End Sub

Private Sub Timer1_Timer()
If Form1.done = True And pb.Value = 240 Then
    Timer1.Enabled = False
    Timer2.Enabled = False
    Timer3.Enabled = True
End If
End Sub

Private Sub Timer2_Timer()
i = Form1.statusbar
If i > 240 Then
    i = 240
End If
pb.Value = i
Label3.Caption = Form1.Status
End Sub

Private Sub Timer3_Timer()
If j = 2 Then
    Call Form1.Show
    Unload Me
End If
j = j + 1
End Sub

