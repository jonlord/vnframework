''' <summary>
''' Allows for easier handling of URL's and their action
''' </summary>
Public Class urlAction
    ''' <summary>
    ''' The base url
    ''' </summary>
    Public url As String
    ''' <summary>
    ''' The complete parameter string
    ''' </summary>
    Public requestString As String
    ''' <summary>
    ''' POST or GET
    ''' </summary>
    Public method As String
    Sub New(url As String, requestString As String, method As String)
        Me.url = url
        Me.requestString = requestString
        Me.method = method
    End Sub
End Class