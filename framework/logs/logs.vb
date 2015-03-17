Public Class logs
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="msg">Message to write</param>
    Public Shared Sub writeToLog(msg As String)
        Debug.Print(msg)
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="msg">Message to write</param>
    Public Shared Sub writeToErrorLog(msg As String)
        Debug.Print(msg)
    End Sub
End Class
