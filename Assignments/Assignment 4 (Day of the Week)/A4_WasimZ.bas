Attribute VB_Name = "Module1"
Option Explicit

Public Sub Initialize(ByVal MaxMonth As Integer, _
                      ByVal MaxDay As Integer, _
                      ByVal MaxYear As Integer, _
                      ByVal StartYear As Integer)
    
    Dim K As Integer
    
    'Adds all month items
    For K = 0 To MaxMonth - 1
        frmMain.cboMonth.AddItem MonthName(K + 1)
    Next K
    
    'Adds all day items
    For K = 0 To MaxDay - 1
        frmMain.cboDay.AddItem K + 1
    Next K
    
    'Adds all year items
    For K = StartYear To StartYear + MaxYear
        frmMain.cboYear.AddItem K
    Next K
    
    'Sets the selected items of the combo boxes of today's date
    'and outputs the corresponding weekday
    frmMain.cboMonth.ListIndex = Month(Now) - 1
    frmMain.cboDay.ListIndex = Day(Now) - 1
    frmMain.cboYear.ListIndex = Year(Now) - 1900
    
    'Calls the find day event procedure for the current date
    frmMain.cmdFindDay_Click

End Sub

Public Function ValidDate(ByVal CDay As Integer, ByVal CMonth As Integer, ByVal CYear As Integer) As Boolean
    
    'The date is default
    ValidDate = True
    
    Select Case CMonth
        'February
        Case 2
            'Checks if February can be 29 if the year is a leap year
            If CDay = 29 Then
                If Not ((CYear Mod 4 = 0) And (CYear Mod 100 <> 0 Or CYear Mod 400 = 0)) Then
                    ValidDate = False
                End If
            ElseIf CDay > 29 Then
                ValidDate = False
            End If
        'April, June, September, November
        'Days with only 30 days
        Case 4, 6, 9, 11
            If CDay > 30 Then
                ValidDate = False
            End If
    End Select
    
End Function

Public Function GetWeekDay(ByVal CMonth As Integer, _
                           ByVal CDay As Integer, _
                           ByVal CYear As Integer) As String
    
    'Temp variables to avoid altering the value of parameters
    Dim Mon As Integer
    Dim Yr As Integer
    
    'Placeholder variables
    Dim W As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim Z As Integer
    
    Dim Century As Integer
    Dim CenturyYear As Integer
    Dim DayNum As Integer
    
    Mon = CMonth
    Yr = CYear
    
    'Zellers algorithm implementation for finding weekdays
    If Mon = 1 Or Mon = 2 Then
        Mon = Mon + 10
        Yr = Yr - 1
    Else
        Mon = Mon - 2
    End If
    
    Century = Yr \ 100
    CenturyYear = Yr Mod 100
    
    W = (13 * Mon - 1) \ 5
    X = CenturyYear \ 4
    Y = Century \ 4
    
    'Adding 7777 to avoid negative values for the WeekDayName function
    Z = W + X + Y + CDay + CenturyYear - 2 * Century + 7777
    
    DayNum = Z Mod 7
    
    'Values from 0 to 6 become 1 to 7 for the WeekDayName function
    GetWeekDay = WeekdayName(DayNum + 1)
    
End Function
