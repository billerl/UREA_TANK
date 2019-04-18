Imports System.Globalization

Public Class Form1
    Public xtraData As String
    Public baseNum As String
    Public tmdNum As String
    Public julYear As String
    Public julWeek As String
    Public julDay As String
    Public serNum As String
    Public seqNum As String
    Public partMatch As Boolean
    Public scanned As String
    Public convMonth As String
    Public convDay As String
    Public calcDayofMonth As String
    Public fisrtThursdayRuleApplies As Boolean
    Public x As Integer



    Private Sub Label14_Click(sender As Object, e As EventArgs) Handles Label14.Click

    End Sub


    Public Sub BreakDown()
        Try

            scanned = TextBox1.Text
            xtraData = scanned.Substring(0, 6)
            Label29.Text = xtraData
            baseNum = scanned.Substring(6, 6)
            Label24.Text = baseNum
            julYear = scanned.Substring(12, 2)
            Label25.Text = julYear
            julWeek = scanned.Substring(14, 2)
            Label26.Text = julWeek
            julDay = scanned.Substring(16, 1)
            Label27.Text = julDay
            seqNum = scanned.Substring(17, 4)
            Label28.Text = seqNum
        Catch ex As Exception
            MessageBox.Show("Serial Number Not Configured Correctly. " & String.Format("Error: {0}", ex.Message))
        End Try

    End Sub
    Public Sub Examine()
        DateConvert()

        If xtraData <> "[)>06T" Or julDay.Length > 1 Or julDay > 7 Or julWeek > 53 Or julWeek.Length <> 2 Or julWeek.Length <> 2 Then
            Label24.ForeColor = Color.Red
            Label25.ForeColor = Color.Red
            Label26.ForeColor = Color.Red
            Label27.ForeColor = Color.Red
            Label28.ForeColor = Color.Red
            Label29.ForeColor = Color.Red
            MessageBox.Show("Not a valid serial number")
            partMatch = False
            Exit Sub
        End If

        If Label6.Text.Contains(baseNum) Then
            Label6.ForeColor = Color.Lime
            Label1.ForeColor = Color.Lime
            Label5.ForeColor = Color.Lime
        ElseIf Label8.Text.Contains(baseNum) Then
            Label8.ForeColor = Color.Lime
            Label2.ForeColor = Color.Lime
            Label7.ForeColor = Color.Lime
        ElseIf Label10.Text.Contains(baseNum) Then
            Label10.ForeColor = Color.Lime
            Label3.ForeColor = Color.Lime
            Label9.ForeColor = Color.Lime
        ElseIf Label12.Text.Contains(baseNum) Then
            Label12.ForeColor = Color.Lime
            Label4.ForeColor = Color.Lime
            Label11.ForeColor = Color.Lime
        ElseIf Label16.Text.Contains(baseNum) Then
            Label16.ForeColor = Color.Lime
            Label4.ForeColor = Color.Lime
            Label11.ForeColor = Color.Lime
        ElseIf Label17.Text.Contains(baseNum) Then
            Label17.ForeColor = Color.Lime
            Label3.ForeColor = Color.Lime
            Label9.ForeColor = Color.Lime
        ElseIf Label18.Text.Contains(baseNum) Then
            Label18.ForeColor = Color.Lime
            Label2.ForeColor = Color.Lime
            Label7.ForeColor = Color.Lime
        ElseIf Label19.Text.Contains(baseNum) Then
            Label19.ForeColor = Color.Lime
            Label1.ForeColor = Color.Lime
            Label5.ForeColor = Color.Lime
        Else
            Label24.ForeColor = Color.Red
            MessageBox.Show("Not a valid serial number. Part number not found")
            partMatch = False
            Exit Sub
        End If
        partMatch = True
    End Sub
    Public Sub DateConvert()
        Select Case julDay
            Case 1
                convDay = "Monday"
            Case 2
                convDay = "Tuesday"
            Case 3
                convDay = "Wednesday"
            Case 4
                convDay = "Thursday"
            Case 5
                convDay = "Friday"
            Case 6
                convDay = "Saturday"
            Case 7
                convDay = "Sunday"
        End Select
        'To find the correct offset(x) for a year, open calander. Starting at the 14th of January 
        'count the days moving backwards until you get to the second week's Sunday
        'Example: 2019 Jan, 14th is on a Monday in the 3rd week so counting the days backwards until
        'you get to 2nd week's Sunday (Jan, 6th) you get the number 8 so 8 is the offset for 2019
        'this program is only configured thru 2026
        'this is where we set up the offset (x)
        Select Case julYear
            Case 17
                x = 6
            Case 18
                x = 7
            Case 19
                x = 8
            Case 20
                x = 9
            Case 21
                x = 4
            Case 22
                x = 5
            Case 23
                x = 6
            Case 24
                x = 7
            Case 25
                x = 9
            Case 26
                x = 10

        End Select
        If julYear = "20" Or julYear = "24" Or julYear = "28" Or julYear = "32" Then
            'Label37.Text = " Leap Year calcDayofMonth is " & julYear & julYear.GetType.ToString
            LeapYear()
        Else
            'Label37.Text = " calcDayofMonth is " & calcDayofMonth
            RegYear()
        End If

    End Sub
    Public Sub RegYear()
        Try
            Dim i As Integer
            i = ((julWeek * 7) - x) + julDay
            'Label46.Text = i
            Select Case i
                Case <= 0
                    convMonth = "December"
                    calcDayofMonth = 31 - i
                    fisrtThursdayRuleApplies = True
                Case 1 To 31
                    convMonth = "January"
                    calcDayofMonth = i - 0
                Case 32 To 59
                    convMonth = "February"
                    calcDayofMonth = i - 31
                Case 60 To 90
                    convMonth = "March"
                    calcDayofMonth = i - 59
                Case 91 To 120
                    convMonth = "April"
                    calcDayofMonth = i - 90
                Case 121 To 151
                    convMonth = "May"
                    calcDayofMonth = i - 120
                Case 152 To 181
                    convMonth = "June"
                    calcDayofMonth = i - 151
                Case 182 To 212
                    convMonth = "July"
                    calcDayofMonth = i - 181
                Case 213 To 243
                    convMonth = "August"
                    calcDayofMonth = i - 212
                Case 244 To 273
                    convMonth = "September"
                    calcDayofMonth = i - 243
                Case 274 To 304
                    convMonth = "October"
                    calcDayofMonth = i - 273
                Case 305 To 334
                    convMonth = "November"
                    calcDayofMonth = i - 304
                Case 335 To 365
                    convMonth = "December"
                    calcDayofMonth = i - 334
                Case >= 366
                    convMonth = "January"
                    fisrtThursdayRuleApplies = True
                    calcDayofMonth = i - 365
            End Select
            'Label37.Text = i & " = " & julWeek & " x 7 - " & x & " + " & julDay & " calcDayofMonth is " & calcDayofMonth
        Catch ex As Exception
            MessageBox.Show("Serial Number Not Configured Correctly. " & String.Format("Error: {0}", ex.Message))
        End Try
    End Sub
    Public Sub LeapYear()
        Try
            Dim i As Integer
            i = ((julWeek * 7) - x) + julDay
            'Label46.Text = i
            Select Case i
                Case <= 0
                    convMonth = "December"
                    calcDayofMonth = 31 + i
                    fisrtThursdayRuleApplies = True
                Case 1 To 31
                    convMonth = "January"
                    calcDayofMonth = i - 0
                Case 32 To 60
                    convMonth = "February"
                    calcDayofMonth = i - 31
                Case 61 To 91
                    convMonth = "March"
                    calcDayofMonth = i - 60
                Case 92 To 121
                    convMonth = "April"
                    calcDayofMonth = i - 91
                Case 122 To 152
                    convMonth = "May"
                    calcDayofMonth = i - 121
                Case 153 To 182
                    convMonth = "June"
                    calcDayofMonth = i - 152
                Case 183 To 213
                    convMonth = "July"
                    calcDayofMonth = i - 182
                Case 214 To 244
                    convMonth = "August"
                    calcDayofMonth = i - 213
                Case 245 To 274
                    convMonth = "September"
                    calcDayofMonth = i - 244
                Case 275 To 305
                    convMonth = "October"
                    calcDayofMonth = i - 274
                Case 306 To 335
                    convMonth = "November"
                    calcDayofMonth = i - 305
                Case 336 To 366
                    convMonth = "December"
                    calcDayofMonth = i - 335
                Case >= 367
                    convMonth = "January"
                    fisrtThursdayRuleApplies = True
                    calcDayofMonth = i - 366

            End Select

        Catch ex As Exception
            MessageBox.Show("Serial Number Not Configured Correctly. " & String.Format("Error: {0}", ex.Message))
        End Try
    End Sub
    Public Shared Function GetISOWeekOfYear(dt As DateTime) As Integer
        Dim cal As Calendar = CultureInfo.InvariantCulture.Calendar
        Dim d As DayOfWeek = cal.GetDayOfWeek(dt)

        If (d >= DayOfWeek.Monday) AndAlso (d <= DayOfWeek.Wednesday) Then
            dt = dt.AddDays(3)
        End If

        Return cal.GetWeekOfYear(dt, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)

    End Function

    Private Sub CompareDate()
        Dim nowWeek As Integer = DatePart(DateInterval.WeekOfYear, Date.Today)
        Dim nowDayString As String = Date.Today.DayOfWeek.ToString()
        Dim nowDayInt As Integer = Date.Today.DayOfWeek
        Dim nowMonth As String = Date.Today.ToLongDateString
        Dim nowYear As String = Date.Today.Year
        nowYear = nowYear.Substring(2, 2)
        Dim nowDate As Integer = DatePart(DateInterval.Day, Date.Today)
        Dim formatDate As String
        Dim weekDif As Integer
        'Label45.Text = GetISOWeekOfYear(dt:=Now) & " " & calcDayofMonth
        If fisrtThursdayRuleApplies Then
            If convMonth = "December" Then
                formatDate = convDay & ", " & convMonth & " " & calcDayofMonth & ", 20" & julYear - 1
            Else
                formatDate = convDay & ", " & convMonth & " " & calcDayofMonth & ", 20" & julYear + 1
            End If
        Else
                formatDate = convDay & ", " & convMonth & " " & calcDayofMonth & ", 20" & julYear
        End If

        Label43.Text = "Today's date is " & vbNewLine & nowMonth & vbNewLine & " The date of scanned label is " & vbNewLine & formatDate & " "
        If nowYear >= 27 Or julYear >= 27 Then
            Label36.ForeColor = Color.Red
            Label36.Text = "THIS PROGRAM IS NOT" & vbNewLine & "VALID PAST THE YEAR 2026"
            Exit Sub
        End If
        If nowYear < julYear Then
            Label36.ForeColor = Color.Red
            Label36.Text = "The date of this label" & vbNewLine & "indicates a future date"
            Label25.ForeColor = Color.Red
            MessageBox.Show("This label was NOT printed today!",
                        "**WARNING** VERIFY DATE")
            Exit Sub
        Else
            If nowYear = julYear Then

                If nowWeek < julWeek Then
                    Label36.ForeColor = Color.Red
                    Label36.Text = "The date of this label" & vbNewLine & "indicates a future date"
                    Label26.ForeColor = Color.Red
                    MessageBox.Show("This label was NOT printed today!",
                        "**WARNING** VERIFY DATE")
                    Exit Sub
                ElseIf nowWeek = julWeek Then
                    If nowDayInt < julDay Then
                        Label36.ForeColor = Color.Red
                        Label36.Text = "The date of this label" & vbNewLine & "indicates a future date"
                        Label27.ForeColor = Color.Red
                        MessageBox.Show("This label was NOT printed today!",
                        "**WARNING** VERIFY DATE")
                        Exit Sub
                    ElseIf nowDayInt > julDay Then
                        Label36.ForeColor = Color.Orange
                        Label36.Text = "This label was printed" & vbNewLine & nowDayInt - julDay & " day(s) ago"
                        Label27.ForeColor = Color.Orange
                        MessageBox.Show("This label was NOT printed today!",
                        "**WARNING** VERIFY DATE")
                        Exit Sub
                    Else
                        Label36.ForeColor = Color.PaleGreen
                        Label36.Text = "This label was printed " & vbNewLine & "today"
                        Exit Sub
                    End If
                Else
                    weekDif = nowWeek - julWeek
                    Label36.ForeColor = Color.Orange
                    Label36.Text = "This label was printed" & vbNewLine & weekDif & " week(s) ago"
                    Label26.ForeColor = Color.Orange
                    MessageBox.Show("This label was NOT printed today!",
                        "**WARNING** VERIFY DATE")
                    Exit Sub
                End If
            ElseIf nowYear < julYear Then
                Label36.ForeColor = Color.Red
                Label36.Text = "The date of this label" & vbNewLine & "indicates a future date"
                Label25.ForeColor = Color.Red
                MessageBox.Show("This label was NOT printed today!",
                        "**WARNING** VERIFY DATE")
                Exit Sub
            ElseIf nowYear >= 27 Or julYear >= 27 Then
                Label36.ForeColor = Color.Red
                Label36.Text = "THIS PROGRAM IS NOT" & vbNewLine & "VALID PAST THE YEAR 2026"
            Else
                Label36.ForeColor = Color.Orange
                Label36.Text = "This label was printed" & vbNewLine & "in 20" & julYear
                Label25.ForeColor = Color.Orange
                MessageBox.Show("This label was NOT printed today!",
                        "**WARNING** VERIFY DATE")
                Exit Sub
            End If
        End If

    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If TextBox1.Text.Length = 21 Then
            DoYoThang()
        Else
            MessageBox.Show("Not a valid serial number. Number of characters should be 21, found " & TextBox1.Text.Length)
        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Reload()
    End Sub
    Public Sub Reload()
        Label6.ForeColor = Color.Silver
        Label1.ForeColor = Color.Silver
        Label2.ForeColor = Color.Silver
        Label3.ForeColor = Color.Silver
        Label4.ForeColor = Color.Silver
        Label5.ForeColor = Color.Silver
        Label6.ForeColor = Color.Silver
        Label7.ForeColor = Color.Silver
        Label8.ForeColor = Color.Silver
        Label9.ForeColor = Color.Silver
        Label10.ForeColor = Color.Silver
        Label11.ForeColor = Color.Silver
        Label12.ForeColor = Color.Silver
        Label16.ForeColor = Color.Silver
        Label17.ForeColor = Color.Silver
        Label18.ForeColor = Color.Silver
        Label19.ForeColor = Color.Silver
        Label24.ForeColor = Color.PaleGreen
        Label25.ForeColor = Color.PaleGreen
        Label26.ForeColor = Color.PaleGreen
        Label27.ForeColor = Color.PaleGreen
        Label28.ForeColor = Color.PaleGreen
        Label29.ForeColor = Color.PaleGreen
        Label24.Text = ""
        Label25.Text = ""
        Label26.Text = ""
        Label27.Text = ""
        Label28.Text = ""
        Label29.Text = ""
        Label36.Text = ""
        Label43.Text = ""
        Label45.Text = ""
        Label46.Text = ""
        xtraData = ""
        baseNum = ""
        tmdNum = ""
        julYear = ""
        julWeek = ""
        julDay = ""
        serNum = ""
        seqNum = ""
        scanned = ""
        convMonth = ""
        convDay = ""
        calcDayofMonth = ""
        x = 0
        TextBox1.Select()
        partMatch = False
        fisrtThursdayRuleApplies = False
        'Label45.Text = GetISOWeekOfYear(dt:=Now) & " " & Now
        For n As Integer = 2017 To 2028
            Console.WriteLine("{0}: week:{1}", n.ToString,
                              GetISOWeekOfYear(New DateTime(n, 12, 31)).ToString)
        Next

    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Reload()
        TextBox1.Text = ""

    End Sub


    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If TextBox1.Text.Length = 21 Then
            DoYoThang()
        Else
            Exit Sub
        End If
    End Sub
    Public Sub DoYoThang()
        Reload()
        BreakDown()
        Examine()
        If partMatch = True Then
            CompareDate()
        Else
            Exit Sub
        End If
    End Sub


End Class
