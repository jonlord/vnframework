Imports System.Windows.Forms

''' <summary>
''' Error Handling manipulation library to facilitate showing errors to the user
''' </summary>
Public Module errorHandling
    Private errorMessage As String = ""

    ''' <summary>
    ''' Get a global error message
    ''' </summary>
    ''' <param name="message">Error Message to save</param>
    Sub setGlobalErrorMessage(message As String)
        errorMessage = message
    End Sub

    ''' <summary>
    ''' Get a global error message
    ''' </summary>
    ''' <returns>Error Message</returns>
    Function getGlobalErrorMessage()
        Return errorMessage
    End Function

    ''' <summary>
    ''' Shows an error to the user, using the best option available between a message box and an error provider
    ''' </summary>
    ''' <param name="message">Message to show. Will be made user friendly if obtained from database</param>
    ''' <param name="myErrorProvider">Error provider</param>
    ''' <param name="userControl">User control associated to the error provider</param>
    Sub showError(ByVal message As String, Optional ByVal myErrorProvider As ErrorProvider = Nothing, Optional ByVal userControl As Control = Nothing)
        message = sqlserver.cleanError(message) 'This can be found in the sql repository; it helps remove technichal info from user errors
        If myErrorProvider Is Nothing Or userControl Is Nothing Then
            MsgBox(message, MsgBoxStyle.Critical, "Error: " & Application.ProductName)
        Else
            myErrorProvider.SetError(userControl, message)
            Try
                userControl.Focus() 'Focus on the user control reporting the error
            Catch ex As Exception
                Debug.WriteLine(ex.Message)
            End Try
        End If
    End Sub
    ''' <summary>
    ''' Shows a message to the user
    ''' </summary>
    ''' <param name="msg">Message to show</param>
    Sub showConfirmation(ByVal msg As String)
        MsgBox(msg, MsgBoxStyle.Information, Application.ProductName)
    End Sub
    ''' <summary>
    ''' Show a Yes/No prompt to the user
    ''' </summary>
    ''' <param name="msg">Message prompt</param>
    ''' <returns>Response the user chose</returns>
    Function showConfirmationYesNo(ByVal msg As String) As MsgBoxResult
        Return MsgBox(msg, MsgBoxStyle.YesNo, Application.ProductName)
    End Function
End Module