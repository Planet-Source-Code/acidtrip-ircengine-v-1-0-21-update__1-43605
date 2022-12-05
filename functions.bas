Attribute VB_Name = "functions"
Type SYSTEMTIME ' 16 Bytes
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Type TIME_ZONE_INFORMATION
    Bias As Long
    StandardName(31) As Integer
    StandardDate As SYSTEMTIME
    StandardBias As Long
    DaylightName(31) As Integer
    DaylightDate As SYSTEMTIME
    DaylightBias As Long
End Type

Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Public Sub SendData(strData As String, sock As CSocket)
    sock.SendData strData & vbCrLf
End Sub
Public Sub AddStatus(Text As String)
    Form1.status.AddItem Text
    Form1.status.Selected(Form1.status.ListCount - 1) = True
End Sub
Public Sub AddText(Text As String)
    Form1.Text1.Text = Form1.Text1.Text & Text & vbNewLine
End Sub
Function GetLocalTZ(Optional ByRef strTZName As String) As Long
Dim objTimeZone As TIME_ZONE_INFORMATION, lngResult&, i&
    
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
Public Function GetUnixTime() As Long
GetUnixTime = Mid(DateDiff("s", Now, "01/01/1970"), 2) - GetLocalTZ
End Function
Public Function sUnixDate(ByVal lValue As Long) As String
    ' Now for the LocalTime function. Take
    '     the long value representing the number
    ' of seconds since January 1, 1970 and c
    '     reate a useable time structure from it.
    ' Return a formatted string YYYY/MM/DD H
    '     H:MM:SS
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

