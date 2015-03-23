Public Class eventLogWriter
    Public Shared Function write(ByVal Entry As String, ByVal AppName As String, Optional ByVal EventType As EventLogEntryType = EventLogEntryType.Information, Optional ByVal LogName As String = "Application") As Boolean
        Dim objEventLog As New EventLog
        Try
            If Not EventLog.SourceExists(AppName) Then EventLog.CreateEventSource(AppName, LogName)
            objEventLog.Source = AppName
            objEventLog.WriteEntry(Entry, EventType)
            Return True
        Catch Ex As Exception
            Return False
        End Try
    End Function
End Class
