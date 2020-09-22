VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "IRCengine Class Sample"
   ClientHeight    =   8115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9165
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   541
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   611
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4875
      Left            =   7020
      TabIndex        =   10
      Top             =   720
      Width           =   2055
      Begin VB.TextBox txtRealName 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   4500
         Width           =   1815
      End
      Begin VB.TextBox txtUserName 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Text            =   "cIRCengine"
         Top             =   3960
         Width           =   1815
      End
      Begin VB.TextBox txtNick 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Text            =   "thresh0ld"
         Top             =   3480
         Width           =   1815
      End
      Begin VB.CommandButton cmdMsg 
         Caption         =   "&Msg"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   25
         Top             =   1500
         Width           =   495
      End
      Begin VB.CheckBox chkPong 
         Caption         =   "Stay Alive"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   2880
         Value           =   1  'Checked
         Width           =   915
      End
      Begin VB.TextBox unixtime 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   270
         Left            =   120
         TabIndex        =   16
         Text            =   "00000000000"
         Top             =   2580
         Width           =   1815
      End
      Begin VB.TextBox txtChan 
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   1140
         Width           =   1815
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&Join"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1500
         Width           =   615
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&Part"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   780
         TabIndex        =   14
         Top             =   1500
         Width           =   615
      End
      Begin VB.ListBox lstChan 
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         IntegralHeight  =   0   'False
         ItemData        =   "frmMain.frx":000C
         Left            =   120
         List            =   "frmMain.frx":000E
         TabIndex        =   11
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Real Name:"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   120
         Left            =   120
         TabIndex        =   31
         Top             =   4320
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "User Name:"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   120
         Left            =   120
         TabIndex        =   28
         Top             =   3780
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nickname:"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   120
         Left            =   120
         TabIndex        =   26
         Top             =   3300
         Width           =   810
      End
      Begin VB.Label txtIdle 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   960
         TabIndex        =   24
         ToolTipText     =   "Idle Time"
         Top             =   2940
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unix Time:"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   120
         Left            =   120
         TabIndex        =   18
         Top             =   2400
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Connection Status:"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   120
         Left            =   120
         TabIndex        =   17
         Top             =   1860
         Width           =   1620
      End
      Begin VB.Label lblState 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "OFFLINE"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2100
         Width           =   1815
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   240
      Top             =   840
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   2430
      Left            =   0
      ScaleHeight     =   162
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   611
      TabIndex        =   1
      Top             =   5685
      Width           =   9165
      Begin VB.Frame Frame2 
         Height          =   675
         Left            =   120
         TabIndex        =   5
         Top             =   420
         Width           =   8955
         Begin VB.TextBox txtPort 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   330
            Left            =   3660
            TabIndex        =   23
            Text            =   "6667"
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtIP 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   330
            Left            =   1440
            TabIndex        =   22
            Text            =   "127.0.0.1"
            Top             =   240
            Width           =   2115
         End
         Begin VB.CommandButton Command5 
            Caption         =   "&Ping User"
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   6
               Charset         =   255
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4620
            TabIndex        =   9
            Top             =   240
            Width           =   1035
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Cl&ose"
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   6
               Charset         =   255
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   7680
            TabIndex        =   8
            Top             =   240
            Width           =   1155
         End
         Begin VB.CommandButton Command3 
            Caption         =   "C&lear Text/Status"
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   6
               Charset         =   255
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5760
            TabIndex        =   7
            Top             =   240
            Width           =   1815
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Connect"
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   6
               Charset         =   255
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.TextBox txtCmd 
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         HideSelection   =   0   'False
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   8955
      End
      Begin VB.Frame Frame1 
         Caption         =   "Parsed Raw Events"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         TabIndex        =   2
         Top             =   1140
         Width           =   8955
         Begin VB.ListBox status 
            BackColor       =   &H00808080&
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   840
            IntegralHeight  =   0   'False
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   8715
         End
      End
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   780
      Width           =   6795
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   0
      ScaleHeight     =   645
      ScaleWidth      =   9135
      TabIndex        =   20
      Top             =   0
      Width           =   9165
      Begin VB.Label lblLogo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IRCengine (Class) Sample"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   660
         Left            =   120
         TabIndex        =   21
         Top             =   -60
         Width           =   6750
      End
   End
   Begin VB.Menu mnuFuck 
      Caption         =   "Fuck"
      Visible         =   0   'False
      Begin VB.Menu cmdHello 
         Caption         =   "Hello"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents irc As IRCengine
Attribute irc.VB_VarHelpID = -1
Dim WithEvents sock As CSocket
Attribute sock.VB_VarHelpID = -1
Dim WithEvents DCCsock As CSocket
Attribute DCCsock.VB_VarHelpID = -1
Dim sockStat As StateConstants
Dim keyHistory As String
Dim lastSent As Long

Private Sub cmdMsg_Click()
Dim Msg As String
If sock.State = sckConnected Then
    Msg = InputBox("Enter Message", "Message to Channel")
    SendData irc.Privmsg(lstChan.List(lstChan.ListIndex), Msg)
End If
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command4_Click()
If sock.State = sckConnected Then
    SendData "QUIT :BYEBYE"
    sock.CloseSocket
End If
End
End Sub

Private Sub Command5_Click()
Dim user As String
If sock.State = sckConnected Then
    user = InputBox("Who to ping?", "Ping Who?")
    sock.SendData "PRIVMSG " & user & " :PING " & GetUnixTime & "" & vbCrLf
End If
End Sub

Private Sub Command6_Click()
If sock.State = sckConnected Then SendData "PART " & lstChan.List(lstChan.ListIndex)
End Sub

Private Sub Command7_Click()
If sock.State = sckConnected Then SendData "JOIN " & txtChan.Text
End Sub

Private Sub DCCsock_OnClose()
    AddStatus "DCC Connection Closed"
End Sub

Private Sub DCCsock_OnConnect()
    AddStatus "You have now establised a DCC Connection"
End Sub

Private Sub DCCsock_OnDataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    DCCsock.GetData strData
    AddStatus "[" & DCCsock.SocketHandle & "] " & "DCC Message: " & strData
End Sub

Private Sub Form_Resize()
On Error Resume Next
Text1.Width = (Form1.ScaleWidth - Frame3.Width) - 17
Text1.Height = (Form1.ScaleHeight - Picture1.ScaleHeight) - 53
Frame3.Left = Text1.Width + 12
Frame3.Height = Text1.Height + 4
End Sub

Private Sub irc_OnAction(Message As String, Channel As String, Nickname As String, UserID As String, UserHost As String, RawMsg As String)
    'This is the /me function in mIRC just incase you're wondering
    AddStatus Nickname & " " & Message
End Sub

Private Sub irc_OnAway(AwayMode As ircModes, RawMsg As String)
    
    Select Case AwayMode
        Case UserAway: AddStatus "You are now Away"
        Case UserBack: AddStatus "You are now Back"
    End Select

End Sub

Private Sub irc_OnBan(Channel As String, BanMask As String, BannedBy As String, UserHost As String, UserID As String, RawMsg As String)
    AddStatus BannedBy & " Banned User with Mask: " & BanMask & " at channel: " & Channel
End Sub

Private Sub irc_OnChanMessage(Message As String, Channel As String, Nickname As String, UserHost As String, UserID As String, RawMsg As String)
    AddStatus "[" & Channel & "]: " & "<" & Nickname & "> " & Message
End Sub

Private Sub irc_OnChannelNames(Names() As String, Total As Integer, Server As String, Channel As String, RawMsg As String)
    AddStatus "There are a total of: " & Total & " users in channel: " & Channel
    
    For i = 0 To UBound(Names)
        AddStatus i + 1 & ") " & Names(i)
    Next
    
End Sub

Private Sub irc_OnCommand(Command As String, Param1 As String, Param2 As String, Nickname As String, UserID As String, UserHost As String, RawMsg As String)
    AddStatus "Command: " & Command & " - " & Param2
End Sub

Private Sub irc_OnCTCPFinger(Nickname As String, UserID As String, UserHost As String, RawMsg As String)
    AddStatus "CTCP FINGER REQUEST from: " & Nickname
End Sub

Private Sub irc_OnCTCPPing(DurationRpl As Integer, PingID As Long, Nickname As String, UserHost As String, UserID As String, RawMsg As String)
    AddStatus "You have been pinged by: " & Nickname & " With PingID: [" & PingID & "]" & " Your Reply for that ping is: " & DurationRpl
    SendData irc.Notice(Nickname, irc.CTCP("PING", CStr(DurationRpl) & " second(s)")) & vbCrLf
End Sub

Private Sub irc_OnCTCPPingReply(Duration As Long, PingID As Long, Nickname As String, UserHost As String, UserID As String, RawMsg As String)
    AddStatus "Ping Reply from: " & Nickname & " is: " & Duration & " (secs)"
End Sub

Private Sub irc_OnCTCPreply(MsgParam As String, Nickname As String, CTCPtype As String, UserID As String, UserHost As String, RawMsg As String)
    AddStatus Nickname & " has replied from your CTCP request: " & CTCPtype & " " & MsgParam
End Sub

Private Sub irc_OnCTCPrequest(Nickname As String, Param As String, CTCPreply As String, UserID As String, UserHost As String, CTCPtype As String, RawMsg As String)
AddStatus "CTCP Request from: " & Nickname & " CTCP Command: " & CTCPtype
End Sub

Private Sub irc_OnCTCPTime(TimeReply As String, Nickname As String, UserID As String, UserHost As String, RawMsg As String)
    SendData irc.Notice(Nickname, "TIME " & Format(Now, "dddd mm/dd/yyyy hh:mm:ss") & "")
End Sub

Private Sub irc_OnCTCPVersion(Nickname As String, UserID As String, UserHost As String, RawMsg As String)
    AddStatus "CTCP VERSION REQUEST from: " & Nickname
    'This is your CTCP Version Reply
    SendData irc.Notice(Nickname, irc.CTCP("VERSION", "Using IRCengine v." & App.Major & "." & App.Minor & "." & App.Revision))
End Sub

Private Sub irc_OnDCCchat(IPadd As String, PortNum As Long, Nickname As String, UserID As String, UserHost As String, RawMsg As String)
    Dim answr As String
    AddStatus "DCC Chat Request from: " & Nickname & " IP: " & IPadd & " at Port: " & PortNum
    answr = MsgBox("Would you like to accept the connection!?", vbYesNo + vbInformation, "DCC Chat Request from: " & Nickname)
    If answr = vbYes Then
        DCCsock.Connect IPadd, PortNum
    End If
End Sub
Private Sub irc_OnDCCsend(IPadd As String, PortNum As Long, FileName As String, FileSize As Long, Nickname As String, UserID As String, UserHost As String, RawMsg As String)
    AddStatus "DCC Send Request from: " & Nickname & " at port: " & PortNum & " Filename: " & FileName & " (" & FileSize & ")"
End Sub

Private Sub irc_OnInvite(Channel As String, Nickname As String, UserHost As String, UserID As String, RawMsg As String)
    AddStatus Nickname & "(" & UserID & ")" & " is inviting you to join " & Channel
End Sub

Private Sub irc_OnKick(Message As String, Channel As String, KickedUser As String, KickedBy As String, UserHost As String, UserID As String, RawMsg As String)
    AddStatus KickedUser & " has been kicked out from " & Channel & " by: " & KickedBy
End Sub

Private Sub irc_OnList(Numeric As Integer, Channel As String, NumOfUsers As Integer, Topic As String, ListType As ircModes, RawMsg As String)
    If ListType = RPL_LISTSTART Then
        AddStatus "START OF LIST"
    ElseIf ListType = RPL_LIST Then
        AddStatus "[" & ListType & "]: " & Channel & " [" & NumOfUsers & "]: " & Topic
    ElseIf ListType = RPL_LISTEND Then
        AddStatus "END OF CHANNEL LIST"
    End If
End Sub

Private Sub irc_OnMode(Channel As String, Nickname As String, flags As String, ModeType As ircModes, ExtraParam As String, UserHost As String, UserID As String, RawMsg As String)
    Select Case ModeType
        Case chanMode: AddStatus "Channel Mode: " & Channel & " set by [" & Nickname & "] Flags: " & "[" & flags & "]" & " Extra Parameters: " & ExtraParam
        Case UserMode
            If Nickname <> Trim$(txtNick) Then txtNick.Text = Nickname
            AddStatus "User Mode: Set by [" & Nickname & "] Flags: " & "[" & flags & "]" & ""
    End Select
End Sub

Private Sub irc_OnMOTD(Numeric As Integer, Message As String, RawMsg As String)
    If Numeric = 375 Then
        AddStatus "START OF MOTD"
    ElseIf Numeric = 376 Then
        AddStatus "END OF MOTD"
    Else
        AddStatus "MOTD: " & Message
    End If
End Sub

Private Sub irc_OnNickChange(OldNick As String, NewNick As String, UserID As String, UserHost As String, RawMsg As String)
    AddStatus OldNick & " has changed nick to --> " & NewNick
    If OldNick = Trim$(txtNick.Text) Then
        txtNick.Text = NewNick
    End If
End Sub

Private Sub irc_OnChannelNotice(Channel As String, Message As String, Nickname As String, UserHost As String, UserID As String, RawMsg As String)
    AddStatus "Channel Notice: " & Message & " from: " & Nickname
End Sub

Private Sub irc_OnNickExist(Numeric As Integer, Nickname As String, RawMsg As String)
    AddStatus "[ERROR]: Nickname is already in use: " & Nickname
End Sub

Private Sub irc_OnPart(Channel As String, Message As String, Nickname As String, UserHost As String, UserID As String, RawMsg As String)

If Message <> "" Then
    AddStatus Nickname & " (" & UserID & ")" & " has part the channel " & Channel & " Message: " & Message
Else
    AddStatus Nickname & " (" & UserID & ")" & " has part the channel " & Channel & " <NO MESSAGE>"
End If
    
If LCase$(Nickname) = LCase$(txtNick) Then
    RemChan Channel
End If

End Sub

Private Sub irc_OnPrivateMessage(Message As String, Nickname As String, UserHost As String, UserID As String, RawMsg As String)
    AddStatus "Private Message From: " & Nickname & " >> " & Message
End Sub

Private Sub irc_OnQuit(Message As String, Nickname As String, UserHost As String, UserID As String, RawMsg As String)
    AddStatus Nickname & " has quit IRC " & "(" & Message & ")"
End Sub

Private Sub irc_OnRawData(RawMsg As String, Param1 As String, Param2 As String)
    AddText ">> " & irc.GetToken(RawMsg, "2-", ":", True)
End Sub

Private Sub irc_OnServerConnect()
    AddStatus "You Are Now Connected!!!!!!!!!!"
    sock.SendData "JOIN #BOO" & vbCrLf
End Sub

Private Sub irc_OnServerError(ErrMessage As String, Param As String, ErrType As errConstants, RawMsg As String)
    Beep
    Select Case ErrType
        Case ERR_NOSUCHNICK
            AddStatus "[" & Param & "]: " & "Cannot find that Nickname dude...Are you looking for a ghost?"
        Case ERR_NICKNAMEINUSE
            txtCmd.SetFocus
            txtCmd.Text = "NICK "
            txtCmd.SelStart = Len(txtCmd.Text)
        Case Else
            AddStatus "[" & ErrType & "] Error: " & "[" & Param & "] " & ErrMessage
    End Select
    
End Sub

Private Sub irc_OnServerNumeric(Numeric As String, Message As String, Server As String, RawMsg As String, NumType As NumericType)
    Select Case Numeric
        Case "433"
            txtCmd.Text = "NICK "
        Case Else
            AddStatus "Numeric: " & "[" & Numeric & "] " & Message
    End Select
End Sub

Private Sub irc_OnServerPing(Server As String, RawMsg As String)
    AddStatus "You have been pinged by server: " & Server
    If chkPong.Value = 1 Then SendData "PONG " & sock.LocalIP & " " & Server
End Sub

Private Sub irc_OnTopic(Channel As String, TopicMsg As String, Nickname As String, UserHost As String, UserID As String, RawMsg As String)
    AddStatus "New Topic in channel: " & Channel & " set by: " & Nickname & " Topic is: " & "[" & TopicMsg & "]"
End Sub

Private Sub irc_OnUnknownData(Data As String)
    AddStatus "Recieved Unknown Data: " & Data
End Sub

Private Sub irc_OnUserNotice(Message As String, Nickname As String, UserHost As String, UserID As String, RawMsg As String)
    AddStatus "User Notice: " & Message & " from: " & Nickname
End Sub

Private Sub irc_OnWhois(Numeric As Integer, Nickname As String, UserName As String, UserHost As String, RealName As String, Param1 As String, Param2 As String, RawMsg As String)
    'Whois Lookup event
   Select Case Numeric
        Case 311
            AddStatus "--------------------------------------------"
            AddStatus ">> Performing Whois Lookup on: " & Nickname & " <<"
            AddStatus "--------------------------------------------"
            AddStatus ">> Real Name: " & RealName
            AddStatus ">> Address: " & UserName & "@" & UserHost
        Case 312
            AddStatus ">> Server Using: " & irc.GetToken(Param1, 3, " ")
        Case 301
            AddStatus ">> " & Nickname & " is on Away Mode, Reason is: " & Param2
        Case 313
            AddStatus ">> " & Nickname & " " & Param2
        Case 317
            AddStatus ">> " & Nickname & " Has Been IDLE for: " & irc.GetToken(Param1, "3", " ") & " seconds"
        Case 319
            AddStatus ">> Channels On: " & Param2
        Case 318
            AddStatus "--------------------------------------------"
            AddStatus "End of Whois"
            AddStatus "--------------------------------------------"
    End Select
    
End Sub

Private Sub Picture1_Resize()
On Error Resume Next
txtCmd.Width = Picture1.ScaleWidth - 15
Frame1.Width = Picture1.ScaleWidth - 15
Frame2.Width = Picture1.ScaleWidth - 15
status.Width = Form1.Width - 655
Command4.Width = (Form1.Width - Command4.Left) - 500
End Sub

Private Sub sock_OnClose()
    AddStatus "Connection Closed by Server"
End Sub

Private Sub sock_OnConnect()
'The RECOMMENDED order for the registration process is:
'1) Password
'2) Nickname
'3) User
'USER <username> <hostname> <servername> <realname>

    'SendData "PASS password"
    SendData "NICK " & txtNick
    SendData "USER " & txtUserName & " " & sock.LocalHostName & " " & sock.RemoteHost & " :" & txtRealName
    SendData "PONG"
End Sub

Private Sub Command1_Click()
If Command1.Caption = "&Connect" Then

    If sock.State <> sckConnected Then
        Text1.Text = ""
        status.Clear
        sock.Connect txtIP.Text, txtPort.Text
    Else
        AddStatus "Dude, you are already connected!"
    End If

    Command1.Caption = "&Disconnect"
Else
    Select Case sock.State
        Case sckConnecting
            sock.CloseSocket
            AddStatus "Connection Aborted"
        Case sckConnected
            sock.SendData "QUIT :BYEBYE" & vbCrLf
            sock.CloseSocket
        Case sckClosed
            Beep
            AddStatus "Dude, you're not online"
        Case sckClosing
            Beep
            AddStatus "Still Closing..."
            sock.CloseSocket
        Case sckError
            Beep
            AddStatus "Sock Error: "
    End Select
    Command1.Caption = "&Connect"
End If
End Sub

Private Sub Command3_Click()
Text1.Text = ""
status.Clear
End Sub

Private Sub Form_Load()
Set irc = New IRCengine
Set sock = New CSocket
Set DCCsock = New CSocket
Set UserInfo = New Collection
Form1.Caption = "IRCengine Class Example (BETA) " & "Version: " & App.Major & "." & App.Minor & "." & App.Revision & " by aCiDtRip"
lblLogo.Caption = lblLogo.Caption & " v." & App.Major & "." & App.Minor & "." & App.Revision
lastSent = irc.unixtime
irc.BlockCmdEvents = False
irc.BlockNumEvents = False
txtRealName.Text = "IRCengine v." & App.Major & "." & App.Minor & "." & App.Revision
End Sub
Public Sub SendData(strData As String)
    sock.SendData strData & vbCrLf
End Sub
Private Sub sock_OnDataArrival(ByVal bytesTotal As Long)
Dim strData As String

sock.GetData strData, vbString, bytesTotal
irc.ProcessData strData

End Sub

Private Sub irc_OnJoin(Channel As String, Nickname As String, UserHost As String, UserID As String, RawMsg As String)
    AddStatus Nickname & " (" & UserID & ")" & " has joined channel " & Channel
    If IsInLst(Channel) = False And Nickname = Trim$(txtNick.Text) Then
        AddChan Channel
    End If
End Sub

Private Sub sock_OnError(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Select Case Number
    Case 10061
        AddStatus "[Error]: " & "Could not connect to server, please check your I.P or Port Address"
    Case Else
        AddStatus "[Error]: " & Description & " (" & Number & ")"
        sock.CloseSocket
End Select
End Sub

Private Sub status_DblClick()
txtCmd.Text = status.List(status.ListIndex)
End Sub

Private Sub Text1_Change()
Text1.SelStart = Len(Text1.Text)
End Sub

Private Sub Timer1_Timer()

Select Case sock.State

    Case sckConnected
        lblState.Caption = "ONLINE"
        lblState.ForeColor = vbGreen
        Command1.Caption = "&Disconnect"
        txtNick.Enabled = False
        txtUserName.Enabled = False
        txtRealName.Enabled = False
    Case sckConnecting
        lblState.Caption = "CONNECTING.."
        lblState.ForeColor = vbYellow
    Case sckClosed, sckError
        lblState.Caption = "OFFLINE"
        lblState.ForeColor = vbRed
        Command1.Caption = "&Connect"
        txtNick.Enabled = True
        txtUserName.Enabled = True
        txtRealName.Enabled = True
    Case sckClosing
        lblState.Caption = "CLOSING.."
        lblState.ForeColor = vbWhite

End Select

unixtime.Text = GetUnixTime
txtIdle.Caption = irc.unixtime - lastSent

End Sub

Private Sub txtCmd_GotFocus()
    With txtCmd
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtCmd_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 38 And keyHistory <> "" Then txtCmd.Text = keyHistory

If KeyCode = 40 Then
    txtCmd.SelStart = 1
    txtCmd.SelLength = Len(txtCmd.Text)
End If

End Sub

Private Sub txtCmd_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 And sock.State = sckConnected Then
    sock.SendData txtCmd.Text & vbCrLf
    keyHistory = txtCmd.Text
    txtCmd.Text = ""
    lastSent = irc.unixtime
    Exit Sub
End If

End Sub

Private Sub txtStatus_Change()
    txtStatus.SelStart = Len(txtStatus.Text)
End Sub

Private Sub AddChan(Chan As String)
If Not IsInLst(Chan) = True Then
    lstChan.AddItem UCase$(Chan)
End If
End Sub
Private Sub RemChan(Chan As String)

For i = 0 To lstChan.ListCount
    If LCase$(lstChan.List(i)) = LCase$(Chan) Then lstChan.RemoveItem i
Next

End Sub

Private Function IsInLst(strLst As String) As Boolean
    
    For i = 0 To lstChan.ListCount
    
        If LCase$(lstChan.List(i)) = LCase$(strLst) Then
            IsInLst = True
            Exit For
        End If
        
    Next
    
End Function

