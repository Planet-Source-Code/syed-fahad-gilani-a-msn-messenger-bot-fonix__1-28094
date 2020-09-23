VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fonix Extensions - Set Away Message"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3855
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   3855
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   300
      Left            =   2760
      TabIndex        =   4
      Top             =   2760
      Width           =   1095
   End
   Begin VB.ComboBox cm 
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   3855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3240
      Top             =   -120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   300
      Left            =   1560
      TabIndex        =   3
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "Form2.frx":0000
      Top             =   600
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Set away  message for:"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SelectedUser As String
Dim LastIndex As Integer

Public Users As IMsgrUsers
Private BlockedUsers As IMsgrUsers

Dim oldusers As Integer
Dim listcount As Integer
Dim X As Integer
Dim everyone As Boolean

Dim goneonce As Integer

Private Sub Command1_Click()
If InStr(1, UCase$(cm.List(cm.ListIndex)), "EVERYONE IN THE LIST") Then

    For X = 0 To Users.Count - 1
        Call Form1.SetMessage(Users.Item(X).LogonName, Text1.Text, X)
    Next X

    everyone = True
End If

If Not everyone Then
    If cm.ListIndex = -1 Then
        Exit Sub
    Else
        Call Form1.SetMessage(Users.Item(cm.ListIndex).LogonName, Text1.Text, cm.ListIndex)
    End If
End If
everyone = False
Form1.Label3.Caption = "Away message has been set. Click to view"
Form1.im1.Visible = False
Form1.im2.Visible = True
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Timer1.Enabled = True
    
    Set Users = Form1.mm.List(MLIST_CONTACT)
    Set BlockedUsers = Form1.mm.List(MLIST_BLOCK)
    
    goneonce = False
    
    
    cm.Clear

    For X = 0 To Users.Count - 1
        cm.AddItem Users.Item(X).FriendlyName
    Next X
    cm.AddItem "Everyone in the list", Users.Count
    
    cm.SelText = "Everyone in the list"
    cm.ListIndex = Users.Count
End Sub

Private Sub Form_Unload(Cancel As Integer)
Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
SelectedUser = cm.SelText
LastIndex = cm.ListIndex


If goneonce = False Then
    oldusers = Form1.oldallusers
End If

Set BlockedUsers = Nothing
Set BlockedUsers = Form1.mm.List(MLIST_BLOCK)

Set Users = Nothing
Set Users = Form1.mm.List(MLIST_CONTACT)



If Form1.counter <> Form1.Counter2 And Form1.doneeverything = True Then

    Form1.doneeverything = False

    cm.Clear

    For X = 0 To Users.Count - 1
        cm.AddItem Users.Item(X).FriendlyName
    Next X
    cm.AddItem "Everyone in the list", Users.Count
    cm.SelText = SelectedUser
    cm.ListIndex = LastIndex
End If

If Form1.goneonce = True Then
    If oldusers <> Users.Count Then
        cm.Clear
    
        For X = 0 To Users.Count - 1
            cm.AddItem Users.Item(X).FriendlyName
        Next X
        cm.AddItem "Everyone in the list", Users.Count
        cm.SelText = SelectedUser
        cm.ListIndex = LastIndex
    End If
End If

goneonce = True
oldusers = Users.Count
End Sub
