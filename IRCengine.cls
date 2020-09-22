VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IRCengine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "IRCengine"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'# ----------------------------------------------------------------------#
'# IRCengine by aCiDtRip, Copyright ©2003                                #
'# ----------------------------------------------------------------------#
'# Oh yes...It is FREE for you to use this crappy code!                  #
'# You should LOOOOVVVEEE MEEE!!!!                                       #
'# Based on RFC 1459 and 2812 (Client Protocol)                          #
'# [Website]: http://www33.brinkster.com/constantanxiety/                #
'# [Email]: acidtrip@most-wanted.com                                     #
'# Do not forget to include WithEvents when you declare the variable     #
'# By the way, some of these codes doesn't belong to me so I wont take   #
'# credit from it.                                                       #
'# Watch out for more version to come!!!!!                               #
'# Soon...I guess (if ever I conquer this laziness inside of me)         #
'# ----------------------------------------------------------------------#

'The following code below (well everything that involves the timezone blabla..
'or whatever you call it) came from vbIRCd
'-------------------------------------------------------------
Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Private Declare Function inet_addr Lib "ws2_32.dll" (ByVal cp As String) As Long

Private Type SYSTEMTIME ' 16 Bytes
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type TIME_ZONE_INFORMATION
    Bias As Long
    StandardName(31) As Integer
    StandardDate As SYSTEMTIME
    StandardBias As Long
    DaylightName(31) As Integer
    DaylightDate As SYSTEMTIME
    DaylightBias As Long
End Type

Private Type myBytes
    B1 As Byte
    B2 As Byte
    B3 As Byte
    B4 As Byte
End Type

Private Type myLong
    Val As Long
End Type

'---------------------------------------------------

'Type Variable for Incoming Data
'We use this to make our lives easier :)
Private Type ParsedData
    strNickname As String 'Contains the Nickname of the Source
    strUserID As String 'Contains the user id of the source
    strHost As String 'Contains the Host of the source
    strEvent As String 'Contains the event type (Either a Numeric or a String of command)
    str1stParam As String 'Contains the First String of the delimeter ":"
    str2ndParam As String 'Contains the Second String of the delimeter ":"
    strRawMessage As String 'Contains the Raw Message
    strOrigin As OriginType 'Contains the Origin Type (Server, Client or Unknown)
End Type
'Numeric Type
Enum NumericType
    rplServerInfo = 0 'Server Info
    rplError = 1 'Error Replies
    rplCommand = 2 'Command Replies
    rplWatchSys = 3 'Watch System for DALnet
    rplUndefined = 4 'Undefined Numeric
End Enum
'Origin Type
Enum OriginType
    O_SERVER = 0
    O_CLIENT = 1
    O_UNKNOWN = 2
End Enum
'IRC Modes
Enum ircModes
    UserAway = 306
    UserBack = 305
    chanMode = 0
    UserMode = 1
    RPL_LISTSTART = 321
    RPL_LIST = 322
    RPL_LISTEND = 323
End Enum
'Error Constants (Based from Jeepster's IRC Numeric Reference)
Enum errConstants
    ERR_NOSUCHNICK = 401
    ERR_NOSUCHSERVER = 402
    ERR_NOSUCHCHANNEL = 403
    ERR_CANNOTSENDTOCHAN = 404
    ERR_TOOMANYCHANNELS = 405
    ERR_WASNOSUCHNICK = 406
    ERR_TOOMANYTARGETS = 407
    ERR_NOSUCHSERVICE = 408
    ERR_NOORIGIN = 409
    ERR_NORECIPIENT = 411
    ERR_NOTEXTTOSEND = 412
    ERR_NOTOPLEVEL = 413
    ERR_WILDTOPLEVEL = 414
    ERR_UNKNOWNCOMMAND = 421
    ERR_NOMOTD = 422
    ERR_NOADMININFO = 423
    ERR_FILEERROR = 424
    ERR_NONICKNAMEGIVEN = 431
    ERR_ERRONEUSNICKNAME = 432
    ERR_NICKNAMEINUSE = 433
    ERR_SERVICENAMEINUSE = 434
    ERR_SERVICECONFUSED = 435
    ERR_NICKCOLLISION = 436
    ERR_BANNICKCHANGE = 437
    ERR_NCHANGETOOFAST = 438
    ERR_TARGETTOFAST = 439
    ERR_SERVICESDOWN = 440
    ERR_USERNOTINCHANNEL = 441
    ERR_NOTONCHANNEL = 441
    ERR_USERONCHANNEL = 443
    ERR_NOLOGIN = 444
    ERR_SUMMONDISABLED = 445
    ERR_USERDISABLED = 446
    ERR_NOTREGISTERED = 451
    ERR_NEEDMOREPARAMS = 461
    ERR_ALREADYREGISTRED = 462
    ERR_NOPERMFORHOST = 463
    ERR_PASSWDMISMATCH = 464
    ERR_YOUREBANNEDCREEP = 465
    ERR_KEYSET = 467
    ERR_ONLYSERVERSCANCHANGE = 468
    ERR_CHANNELISFULL = 471
    ERR_UNKNOWNMODE = 472
    ERR_INVITEONLYCHAN = 473
    ERR_BANNEDFROMCHAN = 474
    ERR_BADCHANNELKEY = 475
    ERR_NEEDREGISTEREDNICK = 477
    ERR_BANLISTFULL = 478
    ERR_NOPRIVILIGES = 481
    ERR_CHANOPRIVSNEEDED = 482
    ERR_CANTKILLSERVER = 483
    ERR_NOOPERHOST = 491
    ERR_NOSERVICEHOST = 492
    ERR_UMODEUNKNOWNFLAG = 501
    ERR_USERSDONTMATCH = 502
    ERR_SILELISTFULL = 511
    ERR_NOSUCHGLINE = 512
    ERR_BADPING = 513
    ERR_LISTSYNTAX = 521
End Enum

'These are the Events for IRCengine
'You could add more if you want to

Public Event OnPrivateMessage(Message As String, Nickname As String, UserHost As String, UserID As String, RawMsg As String)
Public Event OnChanMessage(Message As String, Channel As String, Nickname As String, UserHost As String, UserID As String, RawMsg As String)
Public Event OnJoin(Channel As String, Nickname As String, UserHost As String, UserID As String, RawMsg As String)
Public Event OnPart(Channel As String, Message As String, Nickname As String, UserHost As String, UserID As String, RawMsg As String)
Public Event OnQuit(Message As String, Nickname As String, UserHost As String, UserID As String, RawMsg As String)
Public Event OnChannelNotice(Channel As String, Message As String, Nickname As String, UserHost As String, UserID As String, RawMsg As String)
Public Event OnUserNotice(Message As String, Nickname As String, UserHost As String, UserID As String, RawMsg As String)
Public Event OnInvite(Channel As String, Nickname As String, UserHost As String, UserID As String, RawMsg As String)
Public Event OnMode(Channel As String, Nickname As String, flags As String, ModeType As ircModes, ExtraParam As String, UserHost As String, UserID As String, RawMsg As String)
Public Event OnKick(Message As String, Channel As String, KickedUser As String, KickedBy As String, UserHost As String, UserID As String, RawMsg As String)
Public Event OnBan(Channel As String, BanMask As String, BannedBy As String, UserHost As String, UserID As String, RawMsg As String)
Public Event OnTopic(Channel As String, TopicMsg As String, Nickname As String, UserHost As String, UserID As String, RawMsg As String)
Public Event OnChannelNames(Names() As String, Total As Integer, Server As String, Channel As String, RawMsg As String)
'--------------CTCP Events------------------------------------
Public Event OnCTCPPing(DurationRpl As Integer, PingID As Long, Nickname As String, UserHost As String, UserID As String, RawMsg As String)
Public Event OnCTCPPingReply(Duration As Long, PingID As Long, Nickname As String, UserHost As String, UserID As String, RawMsg As String)
Public Event OnCTCPVersion(Nickname As String, UserID As String, UserHost As String, RawMsg As String)
Public Event OnCTCPFinger(Nickname As String, UserID As String, UserHost As String, RawMsg As String)
Public Event OnCTCPTime(TimeReply As String, Nickname As String, UserID As String, UserHost As String, RawMsg As String)
Public Event OnCTCPrequest(Nickname As String, Param As String, CTCPreply As String, UserID As String, UserHost As String, CTCPtype As String, RawMsg As String)
Public Event OnCTCPreply(MsgParam As String, Nickname As String, CTCPtype As String, UserID As String, UserHost As String, RawMsg As String)
'-------------DCC Events-------------------------------------------
Public Event OnDCCchat(IPadd As String, PortNum As Long, Nickname As String, UserID As String, UserHost As String, RawMsg As String)
Public Event OnDCCsend(IPadd As String, PortNum As Long, FileName As String, FileSize As Long, Nickname As String, UserID As String, UserHost As String, RawMsg As String)
'------------------------------------------------------------------
Public Event OnServerPing(Server As String, RawMsg As String)
Public Event OnServerError(ErrMessage As String, Param As String, ErrType As errConstants, RawMsg As String)
Public Event OnServerNumeric(Numeric As String, Message As String, Server As String, RawMsg As String, NumType As NumericType)
Public Event OnCommand(Command As String, Param1 As String, Param2 As String, Nickname As String, UserID As String, UserHost As String, RawMsg As String)
Public Event OnAction(Message As String, Channel As String, Nickname As String, UserID As String, UserHost As String, RawMsg As String)
Public Event OnUnknownData(Data As String)
'----------------Server Query Events-----------------------------------
Public Event OnMOTD(Numeric As Integer, Message As String, RawMsg As String)
Public Event OnWhois(Numeric As Integer, Nickname As String, UserName As String, UserHost As String, RealName As String, Param1 As String, Param2 As String, RawMsg As String)
'---------------Miscellaneous Features----------------------------------
Public Event OnServerConnect()
Public Event OnList(Numeric As Integer, Channel As String, NumOfUsers As Integer, Topic As String, ListType As ircModes, RawMsg As String)
Public Event OnNickChange(OldNick As String, NewNick As String, UserID As String, UserHost As String, RawMsg As String)
Public Event OnNickExist(Numeric As Integer, Nickname As String, RawMsg As String)
Public Event OnAway(AwayMode As ircModes, RawMsg As String)
Public Event OnRawData(RawMsg As String, Param1 As String, Param2 As String)
'---------------Variables-----------------------------------------------
Private mvarBlockNumEvents As Boolean
Private mvarBlockCmdEvents As Boolean
Private finMOTD As Boolean
Public Property Let BlockNumEvents(ByVal vData As Boolean)
Attribute BlockNumEvents.VB_Description = "Blocks all Incoming Numeric and Command Events"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BlockEvents = 5
    mvarBlockNumEvents = vData
End Property
Public Property Get BlockNumEvents() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BlockEvents
    BlockNumEvents = mvarBlockNumEvents
End Property
Public Property Let BlockCmdEvents(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BlockEvents = 5
    mvarBlockCmdEvents = vData
End Property
Public Property Get BlockCmdEvents() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BlockEvents
    BlockCmdEvents = mvarBlockCmdEvents
End Property
Public Function IP2Long(DottedIP As String) As Variant
    ' errors will result in a zero value
    On Error Resume Next

    Dim i As Byte, pos As Integer
    Dim PrevPos As Integer, num As Integer

    ' string cruncher
    For i = 1 To 4
        ' Parse the position of the dot
        pos = InStr(PrevPos + 1, DottedIP, ".", 1)

        ' If its past the 4th dot then set pos to the last
        'position + 1

        If i = 4 Then pos = Len(DottedIP) + 1

       ' Parse the number from between the dots

        num = Int(Mid(DottedIP, PrevPos + 1, pos - PrevPos - 1))

        ' Set the previous dot position
        PrevPos = pos

        ' No dot value should ever be larger than 255
        ' Technically it is allowed to be over 255 -it just
        ' rolls over e.g.
         '256 => 0 -note the (4 - i) that's the
         'proper exponent for this calculation


      IP2Long = ((num Mod 256) * (256 ^ (4 - i))) + IP2Long

    Next

End Function
Public Function Long2IP(ByVal LongIP As Variant) As String

    On Error GoTo ExitFun

    If LongIP = "" Or LongIP < 0 Then Err.Raise vbObjectError + 1

    Dim i As Integer, num As Currency

    ' big number cruncher
    For i = 1 To 4
        ' break off individual dot values - math out the wazoo
        num = Int(LongIP / 256 ^ (4 - i))

        ' sets up the value for the next calculation
        LongIP = LongIP - (num * 256 ^ (4 - i))

        ' a generic error to flag the exception handler -
        'no dot value should ever be larger than 255
        ' technically it is allowed to be over 255
        ' but it's not possible from this calculation so
        'raise an error
        If num > 255 Then Err.Raise vbObjectError + 1

        ' string builder
        If i = 1 Then
            ' 1st dot value has no leading dot
            Long2IP = num
        Else
            ' other dot values have a leading dot
            Long2IP = Long2IP & "." & num
        End If
    Next

Exit Function
ExitFun:
     Long2IP = "0.0.0.0" '"Invalid Input" ' whatever
End Function
Public Function CTCP(Command As String, Optional Param As String) As String
    
    If IsMissing(Param) Then
        CTCP = Chr(1) & Command & Chr(1)
    Else
        CTCP = Chr(1) & Command & " " & Param & Chr(1)
    End If

End Function
Public Function Notice(Target As String, Message As String) As String
    Notice = "NOTICE " & Target & " :" & Message
End Function
Public Function Privmsg(Target As String, Message As String) As String
    Privmsg = "PRIVMSG " & Target & " :" & Message
End Function
Public Property Get unixtime() As Long
    unixtime = Mid(DateDiff("s", Now, "01/01/1970"), 2) - GetLocalTZ
End Property
Public Function GetLocalTZ(Optional ByRef strTZName As String) As Long
Dim objTimeZone As TIME_ZONE_INFORMATION, lngResult&, i&
    'from vbIRCd
    lngResult = GetTimeZoneInformation&(objTimeZone)
    Select Case lngResult
        Case 0&, 1& 'use standard time
            GetLocalTZ = -(objTimeZone.Bias + objTimeZone.StandardBias) * 60 'into minutes
            For i = 0 To 31
                If objTimeZone.StandardName(i) = 0 Then Exit For
                strTZName = strTZName & Chr(objTimeZone.StandardName(i))
            Next
        
        Case 2& 'use daylight savings time
            GetLocalTZ = -(objTimeZone.Bias + objTimeZone.DaylightBias) * 60  'into minutes
            For i = 0 To 31
                If objTimeZone.DaylightName(i) = 0 Then Exit For
                strTZName = strTZName & Chr(objTimeZone.DaylightName(i))
            Next
    End Select
End Function
Public Function sUnixDate(ByVal lValue As Long) As String
    ' Now for the LocalTime function. Take
    '     the long value representing the number
    ' of seconds since January 1, 1970 and c
    '     reate a useable time structure from it.
    ' Return a formatted string YYYY/MM/DD H
    '     H:MM:SS
    'from vbIRCd
    Dim lSecPerYear, Year&, Month&, Day&, Hour&, Minute&, Second&, Temp&, lSecPerDay, lSecPerHour
    ' [0] = normal year, [1] = leap year
    lSecPerYear = Array(31536000, 31622400)
    lSecPerDay = 86400 ' 60*60*24
    lSecPerHour = 3600 ' 60 * 60
    Year = 70 ' starting point
    ' Calculate the year
    Do While (lValue > 0)
        Temp = isLeapYear(Year)
        If (lValue - lSecPerYear(Temp)) > 0 Then
            lValue = lValue - lSecPerYear(Temp)
            Year = Year + 1
        Else
            Exit Do
        End If
    Loop
    
    'Debug.Print "Year = " & Year
    ' Calculate the month
    Month = 1
    Do While (lValue > 0)
        Temp = secsInMonth(Year, Month)
        If (lValue - Temp) > 0 Then
            lValue = lValue - Temp
            Month = Month + 1
        Else
            Exit Do
        End If
    Loop
    
    'Debug.Print "Month = " & Month
    ' Now calculate day
    Day = 1
    Do While (lValue > 0)
        If (lValue - lSecPerDay) > 0 Then
            lValue = lValue - lSecPerDay
            Day = Day + 1
        Else
            Exit Do
        End If
    Loop
    
    'Debug.Print "Day = " & Day
    ' Now calculate Hour
    Hour = 0
    Do While (lValue > 0)
        If (lValue - lSecPerHour) > 0 Then
            lValue = lValue - lSecPerHour
            Hour = Hour + 1
        Else
            Exit Do
        End If
    Loop
    
    Minute = Int(lValue / 60)
    Second = lValue Mod 60
    Year = Year + 1900
    sUnixDate = Month & "/" & Day & "/" & Year & ", " & Hour & ":" & Minute & ":" & Second
End Function
Private Function isLeapYear(Year As Long) As Integer
    ' Determine if given ANSI datetime struc
    '     t represents a leap year
    ' Private function: assumes valid parame
    '     ters
    Dim nYear%, nIsLeap%
    nYear = Year + 1900


    If ((nYear Mod 4 = 0 And Not (nYear Mod 100)) Or nYear Mod 400 = 0) Then
        nIsLeap = 1 ' its a leap year
    Else
        nIsLeap = 0 ' Not a leap year
    End If
    isLeapYear = nIsLeap
End Function
Private Function secsInMonth(Year As Long, Month As Long) As Long
Dim Taxs As Variant, lResult&, lSecPerMonth
    lSecPerMonth = Array(2678400, 2419200, 2678400, 2592000, _
    2678400, 2592000, 2678400, 2678400, _
    2592000, 2678400, 2592000, 2678400)
    ' Compute result
    lResult = lSecPerMonth(Month - 1)

    If (isLeapYear(Year) And Month = 2) Then lResult = lResult + 86400 ' its February In a leap year
    secsInMonth = lResult
End Function
Private Function secsInYears(Year As Long) As Double
    'from vbIRCd
Dim lResult, thisYear&, Temp&, lSecPerYear
    lResult = 0
    ' 0 = normal year, 1 = leap year
    lSecPerYear = Array(31536000, 31622400)

    If (Year > 97) Then
        ' shorten summation iterations for typic
        '     al cases
        lResult = 883612800 ' seconds To Jan 1,1998 00:00:00
        thisYear = 98
    Else
        ' sum all years since 1970
        thisYear = 70
    End If
    ' Sum total seconds since Jan 1, 1970 00:00:00

    While (thisYear < Year)
        'for ( ; thisYear < year; thisYear++)
        Temp = isLeapYear(thisYear)
        lResult = lResult + lSecPerYear(Temp)
        thisYear = thisYear + 1
    Wend
    secsInYears = lResult
End Function
Public Sub ProcessData(ByVal strData As String)
Attribute ProcessData.VB_UserMemId = 0
    Dim ctr As Long, TmpText() As String
    Dim parsed As ParsedData

    'Splits the Stream of data with vbCrLf as our delimeter
    'Uses the Split function, the data is stored in an array called TmpText

    TmpText = Split(strData, vbCrLf)

    DoEvents

    'The following lines below reads the data line by line and parses
    'each line of text then assigns it to the specified variable type
    
    While ctr < UBound(TmpText)
        'Okay...now lets parse each string of data contained in the TmpText variable line by line
        Call ParseIt(TmpText(ctr), parsed)
        ctr = ctr + 1
    Wend

End Sub

Private Sub ParseIt(ByVal strLine As String, ByRef parsed As ParsedData)

        'The Following steps Parses one line of data at a time (which is strLine)
        'Assigns the values to the specified Properties then
        'determines it's type of event (Command or Numeric)
        'If its neither a Command
        Dim Temp1 As String, Temp2 As String, Temp3 As String

        'Exits the Subroutine if the Line is not a valid IRC message
        'If the line has no leading delimeter ":"

        'Checks if a Server Ping occurs
        'PING :irc.myirc.net
        If UCase$(GetToken(strLine, 1, ":")) = "PING" Then
            RaiseEvent OnServerPing(GetToken(strLine, 2, ":"), strLine)
            Exit Sub
        End If
        
        'Exits the Subroutine if data is an invalid IRC message
        If Left$(strLine, 1) <> ":" Or InStr(1, strLine, ":") = 0 Then
            RaiseEvent OnUnknownData(strLine)
        ElseIf GetToken(GetToken(strLine, 1, ":"), 2, " ") = "" Then
            RaiseEvent OnUnknownData(strLine)
        End If

        'Resets Everything
        parsed.strEvent = ""
        parsed.str2ndParam = ""
        parsed.strHost = ""
        parsed.strNickname = ""
        parsed.strRawMessage = ""
        parsed.strUserID = ""
        parsed.strOrigin = O_SERVER

        'Parses for the First Few Lines
        'We have to get the EVENT first before proceeding with the next steps
        parsed.strEvent = GetToken(GetToken(strLine, 1, ":"), 2, " ")
        'Determines the Origin Type (Client, Server or Unknown)
        parsed.strOrigin = isOrigin(GetToken(GetToken(strLine, 1, ":"), 1, " "))
        'Raw Message
        parsed.strRawMessage = strLine
        
        'Checks if the Line is a Numeric Event or a Command event
        'Then assigns the values to the variables
                
        If IsNumeric(parsed.strEvent) = True And Len(parsed.strEvent) = 3 Then
            'Calls for the Numeric Event
            'Note: Numeric events are only sent by servers and not by clients

            'Let strLine = :irc.myirc.net 250 RAFFY :Total connections 2
            
            'strHost (Acutally the Server Host) returns "irc.myirc.net"
            parsed.strHost = GetToken(GetToken(strLine, 1, ":"), 1, " ")
            'str1stParam returns "irc.myirc.net 250 RAFFY"
            parsed.str1stParam = GetToken(strLine, 1, ":")
            'str2ndParam returns "Total connections 2"
            parsed.str2ndParam = GetToken(strLine, "2-", ":")
            If mvarBlockNumEvents = True Then RaiseEvent OnServerNumeric(parsed.strEvent, parsed.str2ndParam, parsed.strHost, parsed.strRawMessage, GetNumericType(CInt(parsed.strEvent))): Exit Sub
            Call eventNumeric(parsed) 'Calls for the Numeric Event
            RaiseEvent OnRawData(strLine, parsed.str1stParam, parsed.str2ndParam)
        ElseIf Not IsNumeric(parsed.strEvent) Then
            'Calls for the Command Event
            
            'Let strLine = :DOOBiE!breedkillr@127.0.0.1 JOIN :#boo
            
            'str1stParam returns DOOBiE!breedkillr@127.0.0.1 JOIN
            parsed.str1stParam = Trim$(GetToken(strLine, "1", ":"))
            'str2ndParam returns #boo
            parsed.str2ndParam = Trim$(GetToken(strLine, "2-", ":"))

            'Sample Line:
            'With: DOOBiE!breedkillr@127.0.0.1 JOIN
            Temp1 = Trim$(GetToken(parsed.str1stParam, 1, " ")) 'Returns: DOOBiE!breedkillr@127.0.0.1
            Temp2 = Trim$(GetToken(Temp1, 2, "!")) ''Returns: breedkillr@127.0.0.1
            If parsed.strOrigin = O_CLIENT Then
                'Checks if the line is coming from a User
                parsed.strNickname = Trim$(GetToken(Temp1, 1, "!")) 'Returns: DOOBiE (Nickname)
                parsed.strUserID = Trim$(GetToken(Temp2, 1, "@")) 'Returns: breedkillr (UserName/UserID)
                parsed.strHost = Trim$(GetToken(Temp2, 2, "@")) 'Returns: 127.0.0.1 (UserHost)
            ElseIf parsed.strOrigin = O_SERVER Then
                'If it is a server then....
                parsed.strNickname = Temp1
                parsed.strHost = ""
                parsed.strUserID = ""
            Else
                '"Unknown Connection"
                'Assumes it came from the connection
                'from which it was recieved from (RFC 1459)
                parsed.strNickname = Temp1
                parsed.strHost = ""
                parsed.strUserID = ""
            End If
            
            If mvarBlockCmdEvents = True Then RaiseEvent OnCommand(parsed.strEvent, parsed.str1stParam, parsed.str2ndParam, parsed.strNickname, parsed.strUserID, parsed.strHost, parsed.strRawMessage): Exit Sub
            'Calls for the Command Event
            Call eventCmd(parsed)
            RaiseEvent OnRawData(strLine, parsed.str1stParam, parsed.str2ndParam)
        Else
            RaiseEvent OnUnknownData(strLine)
        End If

End Sub
Private Function GetNumericType(Numeric As Integer) As NumericType
'0 = ServerInfo; 1 = Error Reply; 2 = Command Reply; 3 = DALnet WatchSys; 4 = Undefined
    If Numeric >= 1 And Numeric <= 99 Then
        GetNumericType = 0
    ElseIf Numeric >= 200 And Numeric <= 399 Then
        GetNumericType = 2
    ElseIf Numeric >= 400 And Numeric <= 599 Then
        GetNumericType = 1
    ElseIf Numeric > 599 Then
        GetNumericType = 3
    Else
        GetNumericType = 4
    End If
End Function
Private Function isOrigin(strLine As String) As OriginType

    If strLine Like "*!*@*" Then
        'Data is coming from a client
        isOrigin = O_CLIENT
    ElseIf strLine Like "*.*" Then
        'Data is coming from a server
        isOrigin = O_SERVER
    Else
        'Source is Unidentified
        'Automatically assumes that it came
        'from the connection it was recieved from
        isOrigin = O_UNKNOWN
    End If
    
End Function
Public Function GetToken(txtString As String, txtN As Variant, txtDel As String, Optional lstAll As Boolean, Optional stripDel As Boolean) As String
On Error Resume Next
'©aCiDtRip
'The lstAll argument, when set to True, will return the original string if no matching delimeter is found
'the splitDel argument removes all matching delimeters found
Dim endpos As Integer, ctr As Integer, wordNum As Integer
Dim finalString As String, Range As Boolean

If Right$(CStr(txtN), 1) = "-" Then Range = True

FirstStep:
txtN = CInt(Replace(txtN, "-", ""))

If InStr(1, txtString, txtDel, vbTextCompare) = 0 Then
    If lstAll = True Then GetToken = txtString Else: GetToken = ""
    Exit Function
End If

Do
    ctr = ctr + 1
    Text = Mid$(txtString, ctr, 1)
    If Text <> txtDel Then
        Exit Do
    End If
Loop While ctr <> Len(txtString)

NextStep:

wordNum = wordNum + 1 'Increments the Word Count

If endpos = Len(txtString) Then GoTo LastStep

endpos = InStr(ctr + 1, txtString, txtDel, vbTextCompare)

'If it reaches the end of the string, then the length of the string is assigned to the endpos value
If endpos = 0 Then endpos = Len(txtString)

If wordNum = txtN Then
    If Range = True Then finalString = Trim$(Right$(txtString, (Len(txtString) - ctr) + 1)) Else: finalString = Trim$(Replace$(Mid$(txtString, ctr, (endpos - ctr) + 1), txtDel, ""))
Else
    ctr = endpos
    GoTo FirstStep
End If

LastStep:
If finalString = "" Then GetToken = IIf(lstAll = True, txtString, "") Else: GetToken = IIf(stripDel = True, Replace$(finalString, txtDel, ""), finalString)

End Function
Private Sub eventNumeric(ByRef parsed As ParsedData)
        'Triggered when an
        Dim Temp1 As String, Temp2 As String, Temp3 As String
        Dim Temp4 As String, Temp5 As String, Temp6 As String
        Dim lstNames() As String

        Select Case parsed.strEvent
            'Case "001"
                'This is where you could put extra features for the numeric events
            Case "252", "253", "254"
                ':irc.myirc.net 252 thresh0ld 0 :Operator(s) Online
                ':irc.myirc.net 253 thresh0ld 1 :Unknown Connection(s)
                ':irc.myirc.net 254 thresh0ld 0 :Channels Formed
                Temp1 = GetToken(parsed.str1stParam, 4, " ") & " " & parsed.str2ndParam
                RaiseEvent OnServerNumeric(parsed.strEvent, Temp1, parsed.strHost, parsed.strRawMessage, GetNumericType(CInt(parsed.strEvent)))
            Case "321", "322", "323"
                'List Event
                ':irc.myirc.net 321 Raffy Channel :Users  Name
                ':irc.myirc.net 322 Raffy #boo 1 : HELLO GUYS!
                ':irc.myirc.net 323 Raffy :End of /LIST
                If parsed.strEvent = "322" Then
                    Temp1 = GetToken(parsed.str1stParam, 4, " ") 'Channel
                    Temp2 = GetToken(parsed.str1stParam, 5, " ") 'Number of Users
                    RaiseEvent OnList(CInt(parsed.strEvent), Temp1, CInt(Temp2), parsed.str2ndParam, 322, parsed.strRawMessage)
                    Exit Sub
                End If
                RaiseEvent OnList(CInt(parsed.strEvent), Temp1, 0, parsed.str2ndParam, CInt(parsed.strEvent), parsed.strRawMessage)
            Case "375", "372", "376", "422"
                'MOTD Event
                ':irc.myirc.net 375 Raffy :irc.myirc.net Message Of The Day
                ':irc.myirc.net 372 Raffy :- Welcome to the vbIRCd Beta Test Server!
                ':irc.myirc.net 376 Raffy :End of /MOTD Command
                'The OnServerConnect Event is only triggered after the end of a MOTD event and only upon Connection
                'This is done by assigning a local variable called finMOTD (Boolean), so after the user has connected to server
                'the finMOTD variable is set to default value of "FALSE" on the 001 event, then after the MOTD message is done,
                'it is set to True, still don't get it? JUST FIGURE IT OUT ;)
                If parsed.strEvent = "376" And finMOTD = False Then
                    RaiseEvent OnMOTD(CInt(parsed.strEvent), parsed.str2ndParam, parsed.strRawMessage)
                    RaiseEvent OnServerConnect
                    finMOTD = True
                ElseIf parsed.strEvent = "375" Or parsed.strEvent = "372" Then
                    RaiseEvent OnMOTD(CInt(parsed.strEvent), parsed.str2ndParam, parsed.strRawMessage)
                Else
                    RaiseEvent OnMOTD(CInt(parsed.strEvent), parsed.str2ndParam, parsed.strRawMessage)
                End If
            Case "422"
                'MOTD File Missing
                If finMOTD = False Then
                    RaiseEvent OnServerConnect
                    finMOTD = True
                End If
            Case "305", "306"
                'Away Event
                ':irc.myirc.net 306 Raffy :You have been marked as being away
                RaiseEvent OnAway(CInt(parsed.strEvent), parsed.strRawMessage)
            Case "301", "311", "312", "313", "317", "318", "319"
                'Whois Event
                ':screamcore 311 Raffy DOOBiE acidtrip 127.0.0.1 127.0.0.1 :Raffy
                ':screamcore 319 Raffy DOOBiE :@#boo @#tops
                ':screamcore 312 Raffy DOOBiE screamcore :COOL
                ':screamcore 317 Raffy DOOBiE 480 1045702394 :seconds idle
                ':screamcore 313 Raffy DOOBiE :is an IRC operator
                ':screamcore 301 Raffy DOOBiE :BRB
                ':screamcore 318 Raffy DOOBiE :End of /WHOIS list
                Temp1 = GetToken(parsed.str1stParam, 4, " ") 'Nickname
                If parsed.strEvent = "311" Then
                    Temp2 = GetToken(parsed.str1stParam, 5, " ") 'UserName
                    Temp3 = GetToken(parsed.str1stParam, 6, " ") 'UserHost
                    Temp4 = parsed.str2ndParam 'Real Name
                End If
                Temp5 = GetToken(parsed.str1stParam, "3-", " ") '1st Parameter
                Temp6 = parsed.str2ndParam '2nd Parameter
                RaiseEvent OnWhois(CInt(parsed.strEvent), Temp1, Temp2, Temp3, Temp4, Temp5, Temp6, parsed.strRawMessage)
            Case "421"
                'Server Error (Unknown Command)
                Temp1 = GetToken(parsed.str2ndParam, 1, " ") 'Err Command
                Temp2 = GetToken(parsed.str2ndParam, "2-", " ") 'Message
                RaiseEvent OnServerError(Temp2, Temp1, CInt(parsed.strEvent), parsed.strRawMessage)
            Case "433"
                'Nickname In Use Event
                RaiseEvent OnNickExist(ERR_NICKNAMEINUSE, parsed.strNickname, parsed.strRawMessage)
            Case "436"
                MsgBox "NICK COLLISION!"
            Case "353"
                'Names List event
                ':irc.myirc.net 353 Raffy = #boo :Raffy DOOBiE FUCKURASS HAHA FUCKER aCiDtRip`
                Temp1 = GetToken(parsed.str1stParam, 5, " ") 'Channel
                Temp2 = GetToken(parsed.str1stParam, 3, " ") 'Nickname
                lstNames = Split(parsed.str2ndParam, " ")
                Temp3 = UBound(lstNames) 'Total
                RaiseEvent OnChannelNames(lstNames, CInt(Temp3) + 1, parsed.strHost, Temp1, parsed.strRawMessage)
            Case Else
                'Error Replies ranges from 400 to 599
                'Command Replies ranges from 200 to 399
                'Based from RFC 1459 and 2812
                If CInt(parsed.strEvent) >= 400 And CInt(parsed.strEvent) <= 599 Then
                    Temp1 = GetToken(parsed.str1stParam, "4-", " ")
                    RaiseEvent OnServerError(parsed.str2ndParam, Temp1, CInt(parsed.strEvent), parsed.strRawMessage)
                Else
                    If parsed.strEvent = "001" Then finMOTD = False 'Resets the variable for the OnConnect event
                    RaiseEvent OnServerNumeric(parsed.strEvent, parsed.str2ndParam, parsed.strHost, parsed.strRawMessage, GetNumericType(CInt(parsed.strEvent)))
                End If
        End Select
        
End Sub
Private Sub eventCmd(ByRef parsed As ParsedData)
        On Error Resume Next
        Dim Temp1 As String, Temp2 As String, Temp3 As String, Temp4 As String
                
        Select Case parsed.strEvent

            Case "PRIVMSG"

                'Checks first if the Message sent is a valid CTCP Command
                If HasPrefix(parsed.str2ndParam, Chr(1)) And parsed.str2ndParam Like "***" Then
                    Call CTCPevent(parsed)
                    Exit Sub
                End If
                
                Temp1 = GetToken(GetToken(parsed.strRawMessage, 1, ":"), 3, " ") 'Channel/User's Name
                
                'Determine if the Message is Private or Channel
                ':DOOBiE!breedkillr@127.0.0.1 PRIVMSG #boo :hello
                If IsChan(Temp1) = True Then
                    RaiseEvent OnChanMessage(parsed.str2ndParam, Temp1, parsed.strNickname, parsed.strHost, parsed.strUserID, parsed.strRawMessage)
                Else
                    RaiseEvent OnPrivateMessage(parsed.str2ndParam, parsed.strNickname, parsed.strHost, parsed.strUserID, parsed.strRawMessage)
                End If
                
            Case "NOTICE"
            
                'Checks first if the message is a valid CTCP Message
                If HasPrefix(parsed.str2ndParam, Chr(1)) And parsed.str2ndParam Like "***" Then Call CTCPreply(parsed): Exit Sub
                
                ':DOOBiE!breedkillr@127.0.0.mpx-3933 NOTICE Raffy :HELLO Raffy! (Client Notice)
                ':DOOBiE!breedkillr@127.0.0.1 NOTICE #boo :HELLO GUYS! (Channel Notice)
                Temp1 = GetToken(parsed.str1stParam, 3, " ")
                
                'If Temp1 is a channel then....
                If IsChan(Temp1) Then
                    RaiseEvent OnChannelNotice(Temp1, parsed.str2ndParam, parsed.strNickname, parsed.strHost, parsed.strUserID, parsed.strRawMessage)
                Else
                    'If not then we could conclude that it is a UserNotice (or whatever u call that)
                    RaiseEvent OnUserNotice(parsed.str2ndParam, parsed.strNickname, parsed.strHost, parsed.strUserID, parsed.strRawMessage)
                End If
                
            Case "JOIN"
                ':DOOBiE!breedkillr@127.0.0.1 JOIN :#boo
                RaiseEvent OnJoin(parsed.str2ndParam, parsed.strNickname, parsed.strHost, parsed.strUserID, parsed.strRawMessage)
            Case "PART"
                
                If parsed.strRawMessage Like ":*:*:*" Then
                    ':DOOBiE!breedkillr@127.0.0.mpx-3933 PART :#boo :byebye (TYPE1)
                    Temp1 = parsed.str2ndParam 'Channel
                    Temp2 = GetToken(parsed.strRawMessage, "3-", ":") 'Message
                    RaiseEvent OnPart(Temp1, Temp2, parsed.strNickname, parsed.strHost, parsed.strUserID, parsed.strRawMessage)
                ElseIf parsed.strRawMessage Like ":*:*" Then
                    ':WiZ!jto@tolsun.oulu.fi PART #playzone :I lost (TYPE2)
                    ':aCiDtRip`!breedkillr@127.0.0.mpx-3933 PART :#BOO
                    If HasPrefix(GetToken(parsed.str1stParam, 3, " "), "#") = True Then
                        Temp1 = GetToken(parsed.str1stParam, 3, " ") 'Channel
                        Temp2 = parsed.str2ndParam 'Message
                    ElseIf HasPrefix(parsed.str2ndParam, "#") Then
                        Temp1 = parsed.str2ndParam
                        Temp2 = ""
                    End If
                    RaiseEvent OnPart(Temp1, Temp2, parsed.strNickname, parsed.strHost, parsed.strUserID, parsed.strRawMessage)
                ElseIf parsed.strRawMessage Like ":*" Then
                    ':aCiDtRip!acidtrip@127.0.0.1 PART #boo (TYPE3)
                    Temp1 = GetToken(parsed.str1stParam, 3, " ") 'Channel Name
                    RaiseEvent OnPart(Temp1, "", parsed.strNickname, parsed.strHost, parsed.strUserID, parsed.strRawMessage)
                Else
                    'Unknown Data
                    RaiseEvent OnUnknownData(parsed.strRawMessage)
                End If
                
            Case "QUIT"
                ':DOOBiE!breedkillr@127.0.0.1 QUIT :BYE BYE
                RaiseEvent OnQuit(parsed.str2ndParam, parsed.strNickname, parsed.strHost, parsed.strUserID, parsed.strRawMessage)
            Case "NICK"
                ':DOOBiE!breedkillr@127.0.0.1 NICK :WEED (From DOOBiE to WEED)
                Temp1 = parsed.str2ndParam
                If Temp1 = "" Then
                    Temp1 = GetToken(parsed.str1stParam, 3, " ")
                End If
                RaiseEvent OnNickChange(parsed.strNickname, Temp1, parsed.strUserID, parsed.strHost, parsed.strRawMessage)
            Case "INVITE"
                ':WEED!breedkillr@127.0.0.1 INVITE Raffy :#tops <--Weed is inviting Raffy to #tops
                RaiseEvent OnInvite(parsed.str2ndParam, parsed.strNickname, parsed.strHost, parsed.strUserID, parsed.strRawMessage)
            Case "MODE"
                ':DOOBiE!breedkillr@127.0.0.mpx-3933 MODE #boo +oo aCiDtRip` Raffy (Chan Mode)
                ':DOOBiE!breedkillr@127.0.0.mpx-3933 MODE #boo +vv aCiDtRip` Raffy
                Temp1 = GetToken(parsed.str1stParam, 3, " ") 'Channel
                If IsChan(Temp1) = True And Temp1 <> "" Then
                    'ModeType: ChannelMode (0)
                    Temp2 = GetToken(parsed.str1stParam, 4, " ") 'Flags
                    Temp3 = GetToken(parsed.str1stParam, "5-", " ") 'ExtraParameters
                    Temp4 = parsed.strNickname 'Nickname
                    If Temp4 = "" Then Temp4 = GetToken(parsed.str1stParam, 1, " ") 'Nickname
                    RaiseEvent OnMode(Temp1, Temp4, Temp2, 0, Temp3, parsed.strHost, parsed.strUserID, parsed.strRawMessage)
                    ':DOOBiE!breedkillr@127.0.0.mpx-3933 MODE #boo +b *!*@127.0.0.mpx-3933
                    Temp4 = GetToken(parsed.str1stParam, 5, " ") 'BanMask
                    If InStr(1, Temp2, "+b") <> 0 Then RaiseEvent OnBan(Temp1, Temp4, parsed.strNickname, parsed.strHost, parsed.strUserID, parsed.strRawMessage)
                Else
                    'ModeType: UserMode (1)
                    If parsed.str2ndParam <> "" And IsUser(GetToken(parsed.str1stParam, 1, " ")) Then
                        ':thresh0ld!thresh0ld@127.0.0.1 MODE thresh0ld :+i
                        RaiseEvent OnMode("", parsed.strNickname, parsed.str2ndParam, 1, "", parsed.strHost, parsed.strUserID, parsed.strRawMessage)
                    ElseIf parsed.str2ndParam = "" And IsUser(GetToken(parsed.str1stParam, 1, " ")) Then
                        ':thresh0ld!thresh0ld@127.0.0.1 MODE thresh0ld +i
                        RaiseEvent OnMode("", parsed.strNickname, GetToken(parsed.str1stParam, 4, " "), 1, "", parsed.strHost, parsed.strUserID, parsed.strRawMessage)
                    Else
                        ':screamcore MODE thresh0ld +ow (User Mode Set by Server)
                        RaiseEvent OnMode("", GetToken(parsed.str1stParam, 1, " "), GetToken(parsed.str1stParam, 4, " "), 1, "", "", "", parsed.strRawMessage)
                    End If
                End If

            Case "KICK"
                ':DOOBiE!breedkillr@127.0.0.mpx-3933 KICK #boo Raffy :No Reason
                Temp1 = GetToken(parsed.str1stParam, 3, " ") 'Channel
                Temp2 = GetToken(parsed.str1stParam, 4, " ") 'Kicked User
                RaiseEvent OnKick(parsed.str2ndParam, Temp1, Temp2, parsed.strNickname, parsed.strHost, parsed.strUserID, parsed.strRawMessage)
            Case "TOPIC"
                ':DOOBiE!breedkillr@127.0.0.mpx-3933 TOPIC #boo :This is a new topic
                Temp1 = GetToken(parsed.str1stParam, 3, " ") 'Channel
                RaiseEvent OnTopic(Temp1, parsed.str2ndParam, parsed.strNickname, parsed.strHost, parsed.strUserID, parsed.strRawMessage)
            Case Else
                RaiseEvent OnCommand(parsed.strEvent, parsed.str1stParam, parsed.str2ndParam, parsed.strNickname, parsed.strUserID, parsed.strHost, parsed.strRawMessage)
        End Select

End Sub
Private Function GetChan(strText As String, Optional CleanOutChar As String) As String
    
    Dim startPos As Integer, endpos As Integer
        
    startPos = InStr(1, strText, "#", vbTextCompare)

    If startPos = 0 Then Exit Function

    endpos = InStr(startPos + 1, strText, " ")

    If endpos = 0 Then endpos = Len(strText)

    If CleanOutChar <> "" Then
        GetChan = Replace$(Trim$(Mid$(strText, startPos, endpos - startPos)), CleanOutChar, "", , -1)
    Else
        GetChan = Trim$(Mid$(strText, startPos, endpos - startPos))
    End If

End Function
Private Function IsChan(strText As String) As Boolean
    If Left$(strText, 1) = "#" Then IsChan = True Else: IsChan = False
End Function
Private Function IsUser(strText As String) As Boolean
    If strText Like "*!*@*" Then IsUser = True Else: IsUser = False
End Function
Private Function HasPrefix(strText As String, strPrefix As String) As Boolean
    If InStr(1, strText, strPrefix) <> 0 Then HasPrefix = True
End Function

Private Sub CTCPevent(parsed As ParsedData)
'The Subroutine that handles the CTCP Requests
'This is triggered when someone sends out a CTCP Request to you
Dim Command As String
Dim Reply As String
Dim Message As String, Channel As String
Dim Temp1 As String, Temp2 As String, Temp3 As String, Temp4 As String, Temp5 As String

If InStr(1, GetToken(parsed.str2ndParam, 1, Chr(1)), " ") <> 0 Then Command = GetToken(GetToken(parsed.str2ndParam, 1, Chr(1)), 1, " ") Else: Command = GetToken(parsed.str2ndParam, 1, Chr(1))

If Command <> "" Then

    Select Case UCase$(Command)
    'Note: "" = Chr(1)

        Case "PING"
            '':DOOBiE!acidtrip@127.0.0.mpx-3933 PRIVMSG Raffy :PING 1044951875 (Ping Request)
            Temp1 = GetToken(GetToken(parsed.str2ndParam, 1, Chr(1)), 2, " ") 'PingID
            RaiseEvent OnCTCPPing(unixtime - CLng(Temp1), CLng(Val(Temp1)), parsed.strNickname, parsed.strHost, parsed.strUserID, parsed.strRawMessage)
        Case "VERSION"
            RaiseEvent OnCTCPVersion(parsed.strNickname, parsed.strUserID, parsed.strHost, parsed.strRawMessage)
        Case "FINGER"
            RaiseEvent OnCTCPFinger(parsed.strNickname, parsed.strUserID, parsed.strHost, parsed.strRawMessage)
        Case "ACTION"
            ':DOOBiE!breedkillr@127.0.0.1 PRIVMSG #boo :ACTION hello thresh0ld
            Temp1 = GetToken(parsed.str1stParam, 3, " ")
            Temp2 = GetToken(GetToken(parsed.str2ndParam, 1, Chr(1)), "2-", " ")
            RaiseEvent OnAction(Temp2, Temp1, parsed.strNickname, parsed.strUserID, parsed.strHost, parsed.strRawMessage)
        Case "DCC"
            Temp1 = GetToken(parsed.str2ndParam, 1, Chr(1)) 'CTCP Message
            Temp2 = GetToken(Temp1, 2, " ") 'DCCType
            Temp3 = GetToken(Temp1, 4, " ") 'LongIP
            Temp4 = GetToken(Temp1, 5, " ") 'PortNumber
    
            If UCase$(Temp2) = "CHAT" Then
                ':DOOBiE!breedkillr@127.0.0.mpx-3933 PRIVMSG Raffy :DCC CHAT chat 2130706433 1037
                RaiseEvent OnDCCchat(Long2IP(Temp3), CLng(Temp4), parsed.strNickname, parsed.strUserID, parsed.strHost, parsed.strRawMessage)
            ElseIf UCase$(Temp2) = "SEND" Then
                ':DOOBiE!breedkillr@127.0.0.mpx-3933 PRIVMSG thresh0ld :DCC SEND TMP23.$$$ 2130706433 3502 0
                Temp5 = GetToken(Temp1, 3, " ") 'Filename
                Temp6 = GetToken(Temp1, 6, " ") 'Filesize
                RaiseEvent OnDCCsend(Long2IP(Temp3), CLng(Temp4), Temp5, CLng(Temp6), parsed.strNickname, parsed.strUserID, parsed.strHost, parsed.strRawMessage)
            End If
        Case "TIME"
            RaiseEvent OnCTCPTime(unixtime, parsed.strNickname, parsed.strUserID, parsed.strHost, parsed.strRawMessage)
        Case Else
            RaiseEvent OnCTCPrequest(parsed.strNickname, parsed.str2ndParam, Reply, parsed.strUserID, parsed.strHost, Command, parsed.strRawMessage)
    
    End Select

End If

End Sub
Private Sub CTCPreply(parsed As ParsedData)
    'Subroutine that handles different kinds of CTCP Replies
    'This is triggered when a User from a network sends out a Reply from
    'your CTCP request
    Dim Command As String
    Dim Temp1 As String, Temp2 As String, Temp3 As String, Temp4 As String, Temp5 As String
    
    ':DOOBiE!acidtrip@127.0.0.mpx-3933 NOTICE Raffy :VERSION mIRC v6.03 Khaled Mardam-Bey
    ':DOOBiE!acidtrip@127.0.0.mpx-3933 NOTICE Raffy :PING 1044951875
    'Note: Duration = UnixTime - PingID
    Command = GetToken(GetToken(parsed.str2ndParam, 1, Chr(1)), 1, " ")
    
    Select Case UCase$(Command)
        Case "VERSION"
            MsgBox "VERSION REPLY!"
        Case "PING"
            Temp1 = GetToken(GetToken(parsed.str2ndParam, 1, Chr(1)), 2, " ") 'PingID
            Temp2 = unixtime - CLng(Val(Temp1))
            RaiseEvent OnCTCPPingReply(CLng(Val(Temp2)), CLng(Temp1), parsed.strNickname, parsed.strHost, parsed.strUserID, parsed.strRawMessage)
        Case Else
            'Triggered when Other CTCP Replies occur (Undefined CTCP)
            RaiseEvent OnCTCPreply(parsed.str2ndParam, parsed.strNickname, Command, parsed.strUserID, parsed.strHost, parsed.strRawMessage)
    End Select
    
End Sub
