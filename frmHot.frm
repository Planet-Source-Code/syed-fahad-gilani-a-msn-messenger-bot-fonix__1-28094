VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fonix Extensions"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   4815
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHot.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   4815
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Set away &message"
      Height          =   315
      Left            =   2490
      TabIndex        =   6
      Top             =   2745
      Width           =   2280
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Caption         =   "Set yourself &away"
      Height          =   315
      Left            =   60
      MaskColor       =   &H8000000D&
      TabIndex        =   3
      Top             =   2745
      UseMaskColor    =   -1  'True
      Width           =   2280
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4320
      Top             =   0
   End
   Begin MSComctlLib.StatusBar sb 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   3120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2292
            MinWidth        =   2292
            Text            =   "Unread Mail: "
            TextSave        =   "Unread Mail: "
            Object.ToolTipText     =   "Unread Mail"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3245
            MinWidth        =   3245
            Text            =   "Status:"
            TextSave        =   "Status:"
            Object.ToolTipText     =   "Current Messenger Status"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   2716
            MinWidth        =   2716
            TextSave        =   "7:43 AM"
            Object.ToolTipText     =   "Local Time"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000004&
      Caption         =   "Online Users"
      ForeColor       =   &H8000000D&
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   3840
         Top             =   0
      End
      Begin VB.ListBox List1 
         BackColor       =   &H80000001&
         ForeColor       =   &H80000005&
         Height          =   1740
         ItemData        =   "frmHot.frx":030A
         Left            =   120
         List            =   "frmHot.frx":030C
         Sorted          =   -1  'True
         TabIndex        =   2
         ToolTipText     =   "Online Contacts"
         Top             =   240
         Width           =   4575
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   720
         TabIndex        =   5
         Top             =   720
         Width           =   3375
      End
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   6600
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image im2 
      Height          =   480
      Left            =   2325
      Picture         =   "frmHot.frx":030E
      ToolTipText     =   "Click to view away messages"
      Top             =   2235
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label3 
      Caption         =   "No away messages has been set!"
      Height          =   375
      Left            =   2940
      TabIndex        =   7
      ToolTipText     =   "Click to view away messages"
      Top             =   2235
      Width           =   1695
   End
   Begin VB.Image im1 
      Height          =   480
      Left            =   2325
      Picture         =   "frmHot.frx":0750
      ToolTipText     =   "Click to view away messages"
      Top             =   2235
      Width           =   480
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   60
      TabIndex        =   4
      Top             =   2235
      Width           =   1815
   End
   Begin VB.Menu Settings 
      Caption         =   "&Settings"
      Begin VB.Menu authpass 
         Caption         =   "Authorized Password"
      End
   End
   Begin VB.Menu About 
      Caption         =   "&About"
      Begin VB.Menu fonix 
         Caption         =   "About Fonix Extensions"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuSystray 
      Caption         =   "Systray"
      Visible         =   0   'False
      Begin VB.Menu mnuRestore 
         Caption         =   "Restore / Open Fonix"
      End
      Begin VB.Menu Spacer1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Exit Program"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'set reference to  MSN Messenger
Public WithEvents mm As MsgrObject
Attribute mm.VB_VarHelpID = -1
Public Users As IMsgrUsers
Public BlockedUsers As IMsgrUsers
Public allusers As IMsgrUsers
Public UserName As String
Public Users2 As IMsgrUsers
Public BlockedUsers2 As IMsgrUsers
Public AwayReason As String
Public NewLoginName As String
Public Mailcount As Integer
Public done As Boolean
Public ConError As Boolean
Public doneeverything As Boolean
Public Status As String
Public AUTHORIZEDPASSWORD As String
Public errorwasthere As Boolean
Public CurrStatus As Boolean
Public awayMessageString As String

Dim AwayMessage() As String
Dim tempString As String

Public oldblockedusers As Integer
Public oldusers As Integer
Public oldallusers  As Integer

Dim temp As Long
Dim X As Integer
Dim y As Integer
Dim i As Integer
Public counter As Integer
Public Counter2 As Integer
Dim counter3 As Integer
Dim TimeCounter As Integer
Public statusbar As Long

Dim s() As String, max As Integer, current As Integer
Dim textline As String, textline1 As String, textline2 As String

Dim Away As Boolean
Dim DoBack As Boolean
Dim er As Boolean
Dim SelfDestruct As Boolean
Public goneonce  As Boolean

Private Sub authpass_Click()
Form5.Show vbModal
End Sub

Private Sub Command1_Click()
Form2.Show vbModal
End Sub

Private Sub Command2_Click()

If Not Away Then
    Away = True
    Command2.Caption = "Set yourself Back"
    Form1.Caption = "Fonix Extensions - Auto Away"
    sb.Panels.Item(2).Text = "Status: Away"
    Command2.Enabled = False
    Form6.Show
    Form1.Enabled = False
    
    mm.LocalState = MSTATE_AWAY
    While mm.LocalState <> MSTATE_AWAY
        DoEvents
    Wend
    Form1.Enabled = True
    Command2.Enabled = True
    CurrStatus = True
Else
    Away = False
    sb.Panels.Item(2).Text = "Status: Online"
    Command2.Caption = "Set yourself Away"
    Form1.Caption = "Fonix Extensions"
    Command2.Enabled = False
    Form6.Show
    Form1.Enabled = False
    
    mm.LocalState = MSTATE_ONLINE
    
    While mm.LocalState <> MSTATE_ONLINE
        DoEvents
    Wend
    Form1.Enabled = True
    Command2.Enabled = True
    CurrStatus = True
End If
End Sub

Private Sub fonix_Click()
Form4.Show vbModal
End Sub

Private Sub Form_Activate()
If mm.Services.PrimaryService.Status = MSS_NOT_LOGGED_ON Then
    SelfDestruct = True
End If
End Sub

Sub Form_Load()
On Error Resume Next
        
        
        AUTHORIZEDPASSWORD = GetSetting("Fonix Extensions", "Settings", "AuthorizedPassword", "12345")
        
        current = 0
        counter = 0
        Counter2 = 0
        counter3 = 0
        statusbar = 0
        TimeCounter = 3
        SelfDestruct = False
        goneonce = False

        
        Status = "Opening Joke Files"
        statusbar = statusbar + 40
        
        Open_file
        
        Set mm = New MsgrObject
        
        If mm.Services.PrimaryService.Status = MSS_NOT_LOGGED_ON Then
            ConError = True
            SelfDestruct = True
            GoTo ConnectionError
        End If
        
        
        Status = "Creating Messenger Controls"
        statusbar = statusbar + 40
        
        Set Users = mm.List(MLIST_ALLOW)
        Set BlockedUsers = mm.List(MLIST_BLOCK)
        Set allusers = mm.List(MLIST_CONTACT)
        
        er = False
        Away = False
       
        i = 1
                
        Status = "Loading User Lists"
        statusbar = statusbar + 40
        
'==================================================
'                    GET USER LIST
'==================================================

Call ChangeList

'===================================================
'===================================================

        GetInfo
        If counter <> 1 Then
            Label1.Caption = counter & " People online"
        Else
            Label1.Caption = counter & " Person online"
        End If
        
        ReDim AwayMessage(mm.List(MLIST_CONTACT).Count, 1)
        

ConnectionError:
If ConError = True Then
  Form1.Show
  Command2.Enabled = False
  Command1.Enabled = False
  List1.Visible = False
  Label1.Caption = "Not Connected"
  sb.Panels(2).Text = "Status: Not Connected"
  Label2.Caption = "Please switch MSN on before using this program.."
  SelfDestruct = True
  Timer1.Interval = 1000
  Timer2.Enabled = False

  Timer1.Enabled = True
  ConError = False
End If
End Sub


Private Sub GetInfo()
On Error Resume Next

        Status = "Getting Status info"
        statusbar = statusbar + 40
        
        Mailcount = mm.UnreadEmail(0)
        
        If Mailcount > 0 Then
            sb.Panels.Item(1).Text = "Unread Mail: " & Mailcount
        Else
           sb.Panels.Item(1).Text = "Unread Mail: 0"
        End If
        
        Select Case mm.LocalState
            Case MSTATE_ONLINE
            
                sb.Panels.Item(2).Text = "Status: Online"
                Command2.Caption = "Set yourself Away"
                Form1.Caption = "Fonix Extensions"
                Away = False
            
            Case MSTATE_INVISIBLE
                
                Command2.Caption = "Set yourself Away"
                Form1.Caption = "Fonix Extensions"
                Away = False
                sb.Panels.Item(2).Text = "Status: Offline"
                    
            Case MSTATE_BUSY
                        
                Command2.Caption = "Set yourself Away"
                Form1.Caption = "Fonix Extensions"
                mm.LocalState = MSTATE_BUSY
                Away = False
                sb.Panels.Item(2).Text = "Status: Busy"
                        
            Case MSTATE_OUT_TO_LUNCH
            
                Form1.Caption = "Fonix Extensions - Auto Away"
                Command2.Caption = "Set yourself Back"
    
                AwayReason = "I'm having lunch and not available at my desk... talk to you later!"
                Away = True
                sb.Panels.Item(2).Text = "Status: Lunch/Away"
                        
            Case MSTATE_IDLE
                
                Command2.Caption = "Set yourself Back"
                Form1.Caption = "Fonix Extensions - Auto Away"
                AwayReason = "I'm probably not at my desk.. so laters then! Leave a message by the way..."
                Away = True
                sb.Panels.Item(2).Text = "Status: Idle/Away"
                        
                      
            Case MSTATE_AWAY
                
                Form1.Caption = "Fonix Extensions - Auto Away"
                Command2.Caption = "Set yourself Back"
                AwayReason = "I've left for some work, or I'm probably sleeping... so leave a message and I'll get back to you later!"
                Away = True
                sb.Panels.Item(2).Text = "Status: Away"
                
            Case MSTATE_BE_RIGHT_BACK
            
                Command2.Caption = "Set yourself Back"
                Form1.Caption = "Fonix Extensions - Auto Away"
                AwayReason = "Be back in a little while.. wait!"
                Away = True
                sb.Panels.Item(2).Text = "Status: Brb"
        
            Case MSTATE_ON_THE_PHONE
    
                Command2.Caption = "Set yourself Back"
                Form1.Caption = "Fonix Extensions - Auto Away"
                AwayReason = "I'm on the phone.. so I'll finish up and talk to you in a few minutes! Leave a message though."
                Away = True
                sb.Panels.Item(2).Text = "Status: On Phone"
         End Select

done = True
End Sub


Private Sub Form_Unload(Cancel As Integer)
SaveSetting "Fonix Extensions", "Settings", "AuthorizedPassword", AUTHORIZEDPASSWORD

Timer1.Enabled = False
Timer2.Enabled = False

frmSplash.Timer1.Enabled = False
frmSplash.Timer2.Enabled = False
Unload frmSplash

Erase AwayMessage

Set mm = Nothing
Set Users = Nothing
Set BlockedUsers = Nothing
Shell_NotifyIcon NIM_DELETE, nid
Unload Me
End
End Sub

Private Sub im2_Click()
Form3.Show vbModal
End Sub

Private Sub im1_Click()
Form3.Show vbModal
End Sub

Private Sub Label3_Click()
Form3.Show vbModal
End Sub

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
For X = 0 To allusers.Count - 1
    If InStr(1, Trim$(UCase$(List1.List(List1.ListIndex))), Trim$(UCase$(allusers.Item(X).FriendlyName))) Or allusers.Item(X).FriendlyName = List1.List(List1.ListIndex) Then
        MsgBox allusers.Item(X).LogonName, vbInformation, "Fonix Extensions - Buddy Login ID"
    End If
Next X
End Sub

Private Sub mm_OnBuddyPropertyChangeResult(ByVal hr As Long, ByVal pUser As Messenger.IMsgrUser, ByVal ePropType As Messenger.MUSERPROPERTY, ByVal vPropVal As Variant, ByVal pService As Messenger.IMsgrService)
Call ChangeList
End Sub

Private Sub ChangeList()
On Error Resume Next
        List1.Clear
        
        Set Users = mm.List(MLIST_ALLOW)
        Set BlockedUsers = mm.List(MLIST_BLOCK)
        Set allusers = mm.List(MLIST_CONTACT)

        For X = 0 To Users.Count - 1
            If Users.Item(X).State <> MSTATE_OFFLINE And Users.Item(X).State <> MSTATE_UNKNOWN Then
                    List1.AddItem Users.Item(X).FriendlyName
                Counter2 = Counter2 + 1
            End If
        Next X
        
        For X = 0 To BlockedUsers.Count - 1
            If BlockedUsers.Item(X).State <> MSTATE_OFFLINE And BlockedUsers.Item(X).State <> MSTATE_UNKNOWN Then
                    List1.AddItem BlockedUsers.Item(X).FriendlyName & "  (BLOCKED)"
                Counter2 = Counter2 + 1
            End If
        Next X
        
        
        ReDim Preserve AwayMessage(mm.List(MLIST_CONTACT).Count, 1)

        counter = Counter2
        
        doneeverything = True

        If Counter2 <> 1 Then
            Label1.Caption = Counter2 & " People online"
        Else
            Label1.Caption = Counter2 & " Person online"
        End If
        Counter2 = 0
End Sub

Private Sub mm_OnListRemoveResult(ByVal hr As Long, ByVal MLIST As Messenger.MLIST, ByVal pUser As Messenger.IMsgrUser)
If mm.Services.PrimaryService.Status <> 0 Then
    Call ChangeList
End If
End Sub

Private Sub mm_OnListAddResult(ByVal hr As Long, ByVal MLIST As Messenger.MLIST, ByVal pUser As Messenger.IMsgrUser)
If mm.Services.PrimaryService.Status <> 0 Then
    Call ChangeList
End If
End Sub

Private Sub mm_OnLocalPropertyChangeResult(ByVal hr As Long, ByVal ePropType As Messenger.MUSERPROPERTY, ByVal vPropVal As Variant, ByVal pService As Messenger.IMsgrService)
If pService.Status = MSS_LOGGING_OFF Then
    SelfDestruct = True
End If
End Sub

Private Sub mm_OnLocalStateChangeResult(ByVal hr As Long, ByVal mLocalState As Messenger.MSTATE, ByVal pService As Messenger.IMsgrService)
If pService.Status = MSS_LOGGING_OFF Then
    SelfDestruct = True
End If
End Sub

Private Sub mm_OnLogoff()
SelfDestruct = True
End Sub

Private Sub mm_OnLogonResult(ByVal hr As Long, ByVal pService As Messenger.IMsgrService)
On Error Resume Next
    
    If Not Away Then
        sb.Panels.Item(2).Text = "Status: Online"
    End If
            
    If Mailcount > 0 Then
        sb.Panels.Item(1).Text = "Unread Mail: " & Mailcount
    Else
        sb.Panels.Item(1).Text = "Unread Mail: 0"
    End If
End Sub


Private Sub mm_OnTextReceived(ByVal pIMSession As Messenger.IMsgrIMSession, ByVal pSourceUser As Messenger.IMsgrUser, ByVal bstrMsgHeader As String, ByVal bstrMsgText As String, pfEnableDefault As Boolean)
On Error Resume Next

If UCase$(bstrMsgText) = "|JOKE|" Or UCase$(bstrMsgText) = "|JOKES|" Then
    bstrMsgHeader = ""
    bstrMsgHeader = "MIME-Version: 1.0" & vbCrLf & "Content-Type: text/plain; charset=UTF-8" & vbCrLf & "X-MMS-IM-Format: FN=Arial; EF=; CO=800000; CS=0; PF=22" & vbCrLf & vbCrLf
    Call pIMSession.SendText(bstrMsgHeader, Print_ChatupLine, MMSGTYPE_ALL_RESULTS)
    bstrMsgText = ""
    Exit Sub
End If

If UCase$(bstrMsgText) = "|BEEP|" Then
temp = PlaySound(App.Path & "\beep.wav", 0, &H0)
End If

If InStr(1, UCase$(bstrMsgText), "wasssup ?!") Then
    GoTo gettingit
End If

If InStr(1, UCase$(bstrMsgText), "|ONLINEUSERS|") Then
        Dim strTemp As String

        strTemp = Mid$(bstrMsgText, 15, 19)
    
        If strTemp = AUTHORIZEDPASSWORD Then
gettingit:
            Dim j As Integer
            
            j = 0
    
            bstrMsgHeader = ""
            bstrMsgHeader = "MIME-Version: 1.0" & vbCrLf & "Content-Type: text/plain; charset=UTF-8" & vbCrLf & "X-MMS-IM-Format: FN=Arial; EF=; CO=800000; CS=0; PF=22" & vbCrLf & vbCrLf
        
            Set Users = mm.List(MLIST_ALLOW)
            Set BlockedUsers = mm.List(MLIST_BLOCK)
            
            Call pIMSession.SendText(bstrMsgHeader, "----Online Users----", MMSGTYPE_ALL_RESULTS)
            If counter <> 1 Then
                Call pIMSession.SendText(bstrMsgHeader, counter & " Users Currently Online", MMSGTYPE_ALL_RESULTS)
            Else
                Call pIMSession.SendText(bstrMsgHeader, counter & " User Currently Online", MMSGTYPE_ALL_RESULTS)
            End If
            
            For X = 0 To Users.Count - 1
            If Users.Item(X).State <> MSTATE_OFFLINE And Users.Item(X).State <> MSTATE_UNKNOWN Then
                j = j + 1
                Call pIMSession.SendText(bstrMsgHeader, j & ". " & Users.Item(X).FriendlyName, MMSGTYPE_ALL_RESULTS)
            End If
            Next X
            
            For X = 0 To BlockedUsers.Count - 1
            If BlockedUsers.Item(X).State <> MSTATE_OFFLINE And BlockedUsers.Item(X).State <> MSTATE_UNKNOWN Then
                j = j + 1
                Call pIMSession.SendText(bstrMsgHeader, j & ". " & BlockedUsers.Item(X).FriendlyName & "  (BLOCKED)", MMSGTYPE_ALL_RESULTS)
            End If
            Next X
         Call pIMSession.SendText(bstrMsgHeader, "-----------------", MMSGTYPE_ALL_RESULTS)
         bstrMsgText = ""
         Exit Sub
    End If
End If

If UCase$(bstrMsgText) = "|HELP|" Then

    bstrMsgHeader = ""
    bstrMsgHeader = "MIME-Version: 1.0" & vbCrLf & "Content-Type: text/plain; charset=UTF-8" & vbCrLf & "X-MMS-IM-Format: FN=Arial; EF=; CO=800000; CS=0; PF=22" & vbCrLf & vbCrLf
    Call pIMSession.SendText(bstrMsgHeader, "Welcome to Fahad's Fonix Extensions", MMSGTYPE_ALL_RESULTS)
    Call pIMSession.SendText(bstrMsgHeader, "         Main Menu        ", MMSGTYPE_ALL_RESULTS)
    Call pIMSession.SendText(bstrMsgHeader, "--------------------------", MMSGTYPE_ALL_RESULTS)
    Call pIMSession.SendText(bstrMsgHeader, "1) |SetAway| AUTH_PASS                 (Sets the user Away)", MMSGTYPE_ALL_RESULTS)
    Call pIMSession.SendText(bstrMsgHeader, "2) |SetBack| AUTH_PASS                 (Sets the user Back)", MMSGTYPE_ALL_RESULTS)
    Call pIMSession.SendText(bstrMsgHeader, "3) |OnlineUsers| AUTH_PASS           (Sets the user Back)", MMSGTYPE_ALL_RESULTS)
    Call pIMSession.SendText(bstrMsgHeader, "4) |ChangeNick| AUTH_PASS NICKNAME    (Sets a new nick)", MMSGTYPE_ALL_RESULTS)
    Call pIMSession.SendText(bstrMsgHeader, "5) |joke|                            (Displays a random Joke)", MMSGTYPE_ALL_RESULTS)
    Call pIMSession.SendText(bstrMsgHeader, "6) |beep|                            (Beeps the user with a BEEP sound)", MMSGTYPE_ALL_RESULTS)
    Call pIMSession.SendText(bstrMsgHeader, "--------------------------", MMSGTYPE_ALL_RESULTS)
    bstrMsgText = ""
    Exit Sub
End If


bstrMsgHeader = "MIME-Version: 1.0" & vbCrLf & "Content-Type: text/plain; charset=UTF-8" & vbCrLf & "X-MMS-IM-Format: FN=Arial; EF=; CO=0; CS=0; PF=22" & vbCrLf & vbCrLf

If InStr(1, UCase$(bstrMsgText), "|CHANGENICK|") Then
    NewLoginName = Mid$(bstrMsgText, 14, 18)
    
    If InStr(1, NewLoginName, AUTHORIZEDPASSWORD) Then
    
        NewLoginName = Mid$(bstrMsgText, 19, Len(bstrMsgText))
        mm.Services.PrimaryService.FriendlyName = NewLoginName
        Call pIMSession.SendText(bstrMsgHeader, "--- Nickname successfuly changed! ---", MMSGTYPE_ALL_RESULTS)
        bstrMsgText = ""
        bstrMsgHeader = ""
        Exit Sub
    Else
        Call pIMSession.SendText(bstrMsgHeader, "--- Authorization code declined or bad syntax! ---", MMSGTYPE_ALL_RESULTS)
        Exit Sub
    End If
End If

If Not Away Then

    If InStr(1, UCase$(bstrMsgText), "|SETAWAY|") Then
       
        strTemp = Mid$(bstrMsgText, 10, 15)
        
        If Trim$(strTemp) = AUTHORIZEDPASSWORD Then
            Away = False
            Command2_Click
            Call pIMSession.SendText(bstrMsgHeader, "--- User successfuly set to Away mode! ---", MMSGTYPE_ALL_RESULTS)
            bstrMsgText = ""
            bstrMsgHeader = ""
            Exit Sub
        Else
            Call pIMSession.SendText(bstrMsgHeader, "--- Authorization code declined or bad syntax! ---", MMSGTYPE_ALL_RESULTS)
            Exit Sub
        End If
    End If
End If

If Away = True Then

    awayMessageString = bstrMsgText
    
    If Trim$(awayMessageString) <> vbCrLf Then
        DoEvents
    Else
        Exit Sub
    End If

    If InStr(1, UCase$(bstrMsgText), "|SETBACK|") Then
        
        strTemp = Mid$(bstrMsgText, 10, 15)

        If Trim$(strTemp) = AUTHORIZEDPASSWORD Then
            Away = True
            Call pIMSession.SendText(bstrMsgHeader, "--- User successfuly set to Online mode! ---", MMSGTYPE_ALL_RESULTS)
            Command2_Click
        Else
            Call pIMSession.SendText(bstrMsgHeader, "--- Authorization code declined or bad syntax! ---", MMSGTYPE_ALL_RESULTS)
            Exit Sub
        End If
        Exit Sub
    End If
        
    For X = 0 To allusers.Count - 1
        If pSourceUser.LogonName = AwayMessage(X, 0) Then
            Call pIMSession.SendText(bstrMsgHeader, Trim$(AwayMessage(X, 1)), MMSGTYPE_ALL_RESULTS)
            Call pIMSession.SendText(bstrMsgHeader, "(This was an Auto-Message by Fahad's MSN RoBot)", MMSGTYPE_ALL_RESULTS)
            Exit Sub
        End If
    Next X
  bstrMsgText = ""
End If
bstrMsgHeader = ""
End Sub

Private Sub mm_OnUserFriendlyNameChangeResult(ByVal hr As Long, ByVal pUser As Messenger.IMsgrUser, ByVal bstrPrevFriendlyName As String)
If SelfDestruct <> True Then
    Call ChangeList
End If
End Sub

Private Sub mm_OnUserStateChanged(ByVal pUser As Messenger.IMsgrUser, ByVal mPrevState As Messenger.MSTATE, pfEnableDefault As Boolean)
If SelfDestruct <> True Then
    Call ChangeList
End If
End Sub

Private Sub Timer1_Timer()

If SelfDestruct <> True Then
    
    For X = 0 To UBound(AwayMessage)
        If AwayMessage(X, 0) = "" Then
            counter3 = counter3 + 1
        End If
    Next X
    
    If counter3 = UBound(AwayMessage) Then
            Form1.Label3.Caption = "No away messages has been set!"
            Form1.im1.Visible = True
            Form1.im2.Visible = False
    End If

End If

counter3 = 0

If TimeCounter = 0 Then
    Form_Unload (1)
End If

If SelfDestruct = True Then
    Form1.Show
    Command2.Enabled = False
    Command1.Enabled = False
    List1.Visible = False
    Label1.Caption = "Not Connected"
    sb.Panels(2).Text = "Status: Not Connected"
    Label2.Caption = "Please switch MSN on before using this program.."

    Timer2.Enabled = False
    Form1.Caption = "Self Destruction in " & TimeCounter & " Seconds"
    TimeCounter = TimeCounter - 1
Else
    Timer1.Enabled = False
    Timer2.Enabled = True
End If
End Sub

Private Sub Destruct()
    Timer1.Enabled = True
End Sub

Private Sub Timer2_Timer()
On Error Resume Next

        Set Users = mm.List(MLIST_ALLOW)
        Set BlockedUsers = mm.List(MLIST_BLOCK)
        Set allusers = mm.List(MLIST_CONTACT)

        GetInfo
        
        If SelfDestruct = True Then
            Destruct
        End If
        
        counter3 = 0
        
        For X = 0 To UBound(AwayMessage)
            If AwayMessage(X, 0) = "" Then
                counter3 = counter3 + 1
            End If
        Next X
        
        If counter3 = UBound(AwayMessage) + 1 Then
                Form1.Label3.Caption = "No away messages has been set!"
                Form1.im1.Visible = True
                Form1.im2.Visible = False
        End If
        
        counter3 = 0

        
        If mm.Services.PrimaryService.Status = MSS_NOT_LOGGED_ON Then
            ConError = True
            GoTo ConnectionError
        End If
        
        Counter2 = 0
        
        
ConnectionError:
If ConError = True Then
  Form1.Show
  Command2.Enabled = False
  Command1.Enabled = False
  List1.Visible = False
  Label1.Caption = "Not Connected"
  sb.Panels(2).Text = "Status: Not Connected"
  Label2.Caption = "Please switch MSN on before using this program.."
  SelfDestruct = True
  Timer1.Interval = 1000
  Timer1.Enabled = True
  Timer2.Enabled = False
  ConError = False
End If

End Sub

Public Sub SetMessage(logon_name As String, message As String, index As Integer)
AwayMessage(index, 0) = logon_name
AwayMessage(index, 1) = message
End Sub


Public Sub DeleteMessage(logon_name As String)
For X = 0 To allusers.Count - 1
    If InStr(1, Trim$(UCase$(Form3.cm.List(Form3.cm.ListIndex))), Trim$(UCase$(allusers.Item(X).FriendlyName))) Then
        For y = 0 To UBound(AwayMessage)
            If AwayMessage(y, 0) = allusers.Item(X).LogonName Then
                AwayMessage(y, 1) = 0
                AwayMessage(y, 0) = 0
                AwayMessage(y, 1) = ""
                AwayMessage(y, 0) = ""
            End If
        Next y
    Else
    End If
Next X
End Sub


Public Function GetMessage(logon_name As String) As String
For X = 0 To allusers.Count - 1
    If InStr(1, Trim$(UCase$(Form3.cm.List(Form3.cm.ListIndex))), Trim$(UCase$(allusers.Item(X).FriendlyName))) Then
        For y = 0 To UBound(AwayMessage)
            If AwayMessage(y, 0) = allusers.Item(X).LogonName Then
                GetMessage = AwayMessage(y, 1)
            End If
        Next y
    Else
    
    End If
Next X
End Function



'------------------------------------------
'               Jokes Section
'------------------------------------------

Public Sub Open_file()
Status = "Loading file contents"
statusbar = statusbar + 40

X = 0

''CHANGE DATAFILE PATH
Open App.Path & "\" & "blonde.txt" For Input As #1
Do While Not EOF(1)

   Line Input #1, textline1 ' Read line into variable.
   Line Input #1, textline2 ' Read line into variable.
   
   If textline1 <> "" And textline2 <> "" Then
      textline = textline1 & vbCrLf & textline2
      X = X + 1
      ReDim Preserve s(X)
      s(X) = textline
   End If
   
Loop
Close
max = UBound(s)
End Sub


Public Function Random_Num(ByVal max As Integer) As Integer
Randomize Timer
Random_Num = Int((max * Rnd) + 1)
End Function


Public Function Print_ChatupLine() As String
Dim bstrMsgHeader  As String

bstrMsgHeader = "MIME-Version: 1.0" & vbCrLf & "Content-Type: text/plain; charset=UTF-8" & vbCrLf & "X-MMS-IM-Format: FN=Arial; EF=; CO=800000; CS=0; PF=22" & vbCrLf & vbCrLf
X = Random_Num(max)
current = X
Print_ChatupLine = s(X)
End Function

Private Sub mnuexit_Click()
Unload Me
End Sub

Private Sub mnuRestore_Click()
Me.Show
Me.WindowState = vbNormal
Me.Refresh
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim Sys As Long
Sys = X / Screen.TwipsPerPixelX

Select Case Sys
    Case WM_RBUTTONDOWN:
        Me.PopupMenu mnuSystray
        
    Case WM_LBUTTONDBLCLK:
        Me.Show
        WindowState = vbNormal
End Select
End Sub


Private Sub Form_Resize()
If WindowState = vbMinimized Then
Me.Hide
Me.Refresh
With nid
.cbSize = Len(nid)
.hwnd = Me.hwnd
.uId = vbNull
.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
.uCallBackMessage = WM_MOUSEMOVE
.hIcon = Me.Icon
.szTip = Me.Caption & vbNullChar
End With
Shell_NotifyIcon NIM_ADD, nid
Else
    Shell_NotifyIcon NIM_DELETE, nid
End If
End Sub
