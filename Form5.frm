VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Fonix Extensions - View/Change Auth Password"
   ClientHeight    =   510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4710
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form5"
   ScaleHeight     =   510
   ScaleWidth      =   4710
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Default         =   -1  'True
      Height          =   315
      Left            =   3720
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Set"
      Height          =   315
      Left            =   2760
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Authorized Password:"
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
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
If Len(Text1.Text) <> 5 Then
    MsgBox "Please insert a 5 digit code!"
    Exit Sub
End If
SaveSetting "Fonix Extensions", "Settings", "AuthorizedPassword", Text1.Text
Form1.AUTHORIZEDPASSWORD = Text1.Text
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Text1.Text = GetSetting("Fonix Extensions", "Settings", "AuthorizedPassword", "12345")
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Then
    Text1.Text = Text1.Text & Chr$(KeyAscii)
End If
End Sub

Private Sub form_load()
Text1.MaxLength = 5
End Sub

Private Sub Text1_Click()
Text1.Text = ""
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
