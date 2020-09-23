VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fonix Extensions - View Away messages"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   4095
   StartUpPosition =   3  'Windows Default
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
      TabIndex        =   4
      Top             =   240
      Width           =   4095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Cancel"
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
      Left            =   2760
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3360
      Top             =   0
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Ok"
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
      Left            =   1440
      TabIndex        =   2
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Delete"
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
      Top             =   2520
      Width           =   1215
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
      Height          =   1815
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   600
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Check away message for:"
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
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "Form3"
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

Dim goneonce  As Boolean

Private Sub cm_Click()
If cm.ListIndex = -1 Then
    Exit Sub
End If

If InStr(1, cm.List(cm.ListIndex), "Everyone") Then
    Text1.Text = ""
    Text1.Text = Form1.GetMessage(Users.Item(0).LogonName)
Else
    Text1.Text = ""
    Text1.Text = Form1.GetMessage(Users.Item(cm.ListIndex).LogonName)
End If
End Sub

Private Sub cm_DblClick()
For X = 0 To Users.Count - 1
    If Users.Item(X).FriendlyName = cm.List(cm.ListIndex) Then
        MsgBox Users.Item(X).LogonName, vbInformation, "Fonix Extensions - Buddy Login ID"
    End If
Next X
End Sub

Private Sub Command1_Click()
Form1.DeleteMessage (Users.Item(cm.ListIndex).LogonName)
Text1.Text = ""
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
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
    cm.SelText = Users.Item(0).FriendlyName
    cm.ListIndex = 0
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
    
    cm.SelText = SelectedUser
    cm.ListIndex = LastIndex
    
End If

If Form1.goneonce = True Then
    If oldusers <> Users.Count Then
        cm.Clear
    
        For X = 0 To Users.Count - 1
            cm.AddItem Users.Item(X).FriendlyName
        Next X
        cm.SelText = SelectedUser
        cm.ListIndex = LastIndex
    End If
End If

goneonce = True
oldusers = Users.Count
End Sub
