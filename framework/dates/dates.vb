Imports System.Windows.Forms
Public Module dates
    ''' <summary>
    ''' Gets a month name by its number
    ''' </summary>
    ''' <param name="index">Number of Month</param>
    ''' <returns>Month Name</returns>
    Public Function monthName(ByVal index As Integer) As String
        Select Case index
            Case 1
                Return JAN
            Case 2
                Return FEB
            Case 3
                Return MAR
            Case 4
                Return APR
            Case 5
                Return MAY
            Case 6
                Return JUN
            Case 7
                Return JUL
            Case 8
                Return AUG
            Case 9
                Return SEP
            Case 10
                Return OCT
            Case 11
                Return NOV
            Case 12
                Return DEC
        End Select
        Return ""
    End Function
    ''' <summary>
    ''' Gets a month name of the previous month by its number
    ''' </summary>
    ''' <param name="index">Number of month</param>
    ''' <returns>Month Name</returns>
    Public Function prevMonthName(index As Integer) As String
        index -= 1
        If index = 0 Then index = 12
        Return monthName(index)
    End Function
    ''' <summary>
    ''' Gets a month index by its name
    ''' </summary>
    ''' <param name="monthName">Month Name</param>
    ''' <returns>Month Name</returns>
    Public Function monthIndex(ByVal monthName As String) As Integer
        Select Case LCase(monthName)
            Case LCase(JAN)
                Return 1
            Case LCase(FEB)
                Return 2
            Case LCase(MAR)
                Return 3
            Case LCase(APR)
                Return 4
            Case LCase(MAY)
                Return 5
            Case LCase(JUN)
                Return 6
            Case LCase(JUL)
                Return 7
            Case LCase(AUG)
                Return 8
            Case LCase(SEP)
                Return 9
            Case LCase(OCT)
                Return 10
            Case LCase(NOV)
                Return 11
            Case LCase(DEC)
                Return 12
        End Select
        Throw New NotImplementedException()
    End Function
    ''' <summary>
    ''' Gets the amount of days in a month
    ''' </summary>
    ''' <param name="monthName">Month Name</param>
    ''' <returns>Amount of days in month</returns>
    Public Function daysInMonth(year As Integer, monthName As String) As Integer
        Dim month As Integer = monthIndex(monthName)
        Return DateTime.DaysInMonth(year, month)
    End Function

    ''' <summary>
    ''' Extracts the date as a string from a DateTimePicker
    ''' </summary>
    ''' <param name="dateTimePicker">DateTimePicker to use</param>
    ''' <returns>The date as a string</returns>
    Public Function extractDateAsString(ByVal dateTimePicker As DateTimePicker) As String
        Return extractDateAsString(dateTimePicker.Value)
    End Function
    ''' <summary>
    ''' Extracts the time as a string from a Date
    ''' </summary>
    ''' <param name="myDate">The date to use</param>
    ''' <returns>The time as a string</returns>
    Public Function extractTimeAsString(ByVal myDate As Date) As String
        Return myDate.Hour & ":" & myDate.Minute & ":" & myDate.Second
    End Function
    ''' <summary>
    '''  Extracts the date as a string from a DateTimePicker
    ''' </summary>
    ''' <param name="myDate">The date to use</param>
    ''' <returns>The date as a string</returns>
    Public Function extractDateAsString(ByVal myDate As Date) As String
        Return myDate.Month & "/" & myDate.Day & "/" & myDate.Year
    End Function

    ''' <summary>
    ''' Get Diffrence in seconds between two dates
    ''' </summary>
    ''' <param name="date1">Fist Date To Compare</param>
    ''' <param name="date2">Second Date To Compare</param>
    ''' <returns>Diffrence in Seconds</returns>
    Public Function getDiffSeconds(date1 As Date, date2 As Date) As Long
        getDiffSeconds = Math.Round((date2 - date1).TotalSeconds * 24 * 60 * 60, 0)
    End Function
End Module
