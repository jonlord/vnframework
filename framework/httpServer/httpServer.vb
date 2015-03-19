Imports System.Reflection
Imports System.Net

Public Class httpServer
    Private Const numRequestsToBeHandled As Integer = 10
    Private nameSpaceName As String
    Sub Main(nameSpaceName As String, Optional port As Integer = 8080, Optional urlListener As String = "listenGET")
        Dim prefixes(0) As String
        prefixes(0) = String.Format("http://*:{0}/{1}/", port, urlListener)
        ProcessRequests(prefixes)
    End Sub

    Private Shared Function fetchInstance(ByVal fullyQualifiedClassName As String) As Object
        Dim o As Object = Nothing
        Try
            o = Activator.CreateInstance(Type.GetType(fullyQualifiedClassName))
        Catch ex As Exception
            logs.writeToErrorLog(ex.Message)
        End Try
        Return o
    End Function

    Private Sub ProcessRequests(ByVal prefixes() As String)
        If Not HttpListener.IsSupported Then
            Console.WriteLine(HTTPSYSTEMREQUIREMENTS)
            Exit Sub
        End If

        ' URI prefixes are required,
        If prefixes Is Nothing OrElse prefixes.Length = 0 Then
            Throw New ArgumentException("Prefixes not defined")
        End If

        ' Create a listener and add the prefixes.
        Dim listener As System.Net.HttpListener = New System.Net.HttpListener()

        For Each s As String In prefixes
            listener.Prefixes.Add(s)
        Next

        Try
            ' Start the listener to begin listening for requests.
            listener.Start()
            Console.WriteLine("Listening...")

            ' Set the number of requests this application will handle.
            For i As Integer = 0 To numRequestsToBeHandled
                Dim response As HttpListenerResponse = Nothing
                Try
                    ' Note: GetContext blocks while waiting for a request.
                    Dim context As HttpListenerContext = listener.GetContext()
                    Dim request = context.Request.QueryString

                    'Dim action As String = reques.Item("action")
                    Dim parameters As New Dictionary(Of String, String)
                    For Each r In request.Keys
                        parameters.Add(r, request.Item(r))
                    Next
                    Dim functionResponse As String = callFromHTTP(parameters)
                    ' Create the response.
                    response = context.Response
                    Dim responseString As String = String.Format("<HTML><meta http-equiv='Content-Type' content='text/html; charset=utf-8'><BODY><h2>{0}</h2></BODY></HTML>", functionResponse)
                    Dim buffer() As Byte =
                        Text.Encoding.UTF8.GetBytes(responseString)
                    response.ContentLength64 = buffer.Length
                    Dim output As System.IO.Stream = response.OutputStream
                    output.Write(buffer, 0, buffer.Length)

                Catch ex As HttpListenerException
                    Console.WriteLine(ex.Message)
                Finally
                    If response IsNot Nothing Then
                        response.Close()
                    End If
                End Try
            Next
        Catch ex As HttpListenerException
            Console.WriteLine(ex.Message)
        Finally
            ' Stop listening for requests.
            listener.Close()
            Console.WriteLine("Done Listening...")
        End Try
    End Sub

    Function callFromHTTP(parameters As Dictionary(Of String, String)) As String
        Dim moduleName As String = nameSpaceName & "." & parameters("module")
        Dim genericObject As Object = fetchInstance(moduleName)
        Dim method As MethodInfo = genericObject.GetType().GetMethod(parameters("action"))
        If Not method Is Nothing Then
            Dim arrParameters(0)
            arrParameters(0) = parameters
            Return method.Invoke(genericObject, arrParameters)
        Else
            Return HTTPMETHODNOTFOUND
        End If
    End Function
End Class