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
End Module
