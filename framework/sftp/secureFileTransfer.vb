Imports Tamir.SharpSsh

''' <summary>
''' Wrapper class for using SFTP, implementing the Tamir.SharpSSH library
''' </summary>
Public Class secureFileTransfer
    Private WithEvents transfer As Sftp
    Private _hostname As String
    Private _port As Integer
    Private _username As String
    Private _password As String
    Private _identityFile As String

    ''' <summary>
    ''' Constucts an SFTP connection using username / password
    ''' </summary>
    ''' <param name="hostName">SFTP Server Hostname</param>
    ''' <param name="port">SFTP Server port</param>
    ''' <param name="userName">SFTP Server username</param>
    ''' <param name="password">SFTP Server password</param>
    Public Sub New(ByVal hostName As String, ByVal port As Integer, ByVal userName As String, ByVal password As String)
        _hostname = hostName
        _port = port
        _username = userName
        _password = password
    End Sub
    ''' <summary>
    ''' Constucts an SFTP connection using username / identityFile
    ''' </summary>
    ''' <param name="hostName">SFTP Server Hostname</param>
    ''' <param name="userName">SFTP Server username</param>
    ''' <param name="identityFile">Path to identity file</param>
    ''' <param name="port">SFTP Server port</param>
    Public Sub New(ByVal hostName As String, ByVal userName As String, identityFile As String, ByVal port As Integer)
        _hostname = hostName
        _port = port
        _username = userName
        _identityFile = identityFile
    End Sub
    ''' <summary>
    ''' Retrieves a file
    ''' </summary>
    ''' <param name="remotePath">Remote path to get from</param>
    ''' <param name="localFile">Local path to store to</param>
    ''' <returns>True if successfull, false if it fails</returns>
    Public Function getFile(ByVal remotePath As String, ByVal localFile As String) As Boolean
        Try
            Try
                If _password <> "" Then
                    transfer = New Sftp(Me._hostname, Me._username, _password)
                ElseIf _password = "" And files.exists(_identityFile) Then
                    transfer = New Sftp(Me._hostname, Me._username)
                    transfer.AddIdentityFile(_identityFile)
                End If
                transfer.Connect(Me._port)
                transfer.Get(remotePath, localFile)
                transfer.Close()
                Return True
            Catch ex As Tamir.SharpSsh.jsch.SftpException
                Debug.Print(ERRORDOWNLOADINGFILE, ex.message)
                Return False
            End Try
        Catch ex2 As Exception
            Debug.Print(ERRORDOWNLOADINGFILE, ex2.Message)
            Return False
        End Try
    End Function
    ''' <summary>
    ''' Uploads a file to the server
    ''' </summary>
    ''' <param name="remotePath">Remote path to store</param>
    ''' <param name="localFile">Local file to upload to</param>
    ''' <returns>True if successfull, false if it fails</returns>
    Public Function putFile(ByVal localFile As String, ByVal remotePath As String) As Boolean
        Try
            Try
                If _password <> "" Then transfer = New Sftp(Me._hostname, Me._username, _password)
                If _password = "" Then
                    transfer = New Sftp(Me._hostname, Me._username)
                    transfer.AddIdentityFile(_identityFile)
                End If

                transfer.Connect(Me._port)
                Try
                    transfer.Mkdir(remotePath)
                Catch
                End Try
                transfer.Put(localFile, remotePath)
                transfer.Close()
                Return True
            Catch ex As Tamir.SharpSsh.jsch.SftpException
                Debug.Print(ERRORUPLADINGFILE, ex.message)
                Return False
            End Try
        Catch ex2 As Exception
            Debug.Print(ERRORUPLADINGFILE, ex2.Message)
            Return False
        End Try
    End Function
    ''' <summary>
    ''' Retrieves the directory list
    ''' </summary>
    ''' <param name="remotePath">Folder to retreive from</param>
    ''' <returns>ArrayList of all files</returns>
    Public Function getDirectoryList(ByVal remotePath As String) As ArrayList
        Dim response As New ArrayList
        Try
            If _password <> "" Then transfer = New Sftp(Me._hostname, Me._username, _password)
            If _password = "" Then
                transfer = New Sftp(Me._hostname, Me._username)
                transfer.AddIdentityFile(_identityFile)
            End If

            transfer.Connect(_port)
            response = transfer.GetFileList(remotePath)
            transfer.Close()
        Catch ex As Exception
            Debug.Print(ERRORGETDIRECTORYLIST, ex.ToString)
        End Try
        Return response
    End Function

    Private Event onTransferStart(ByVal source As String, ByVal destination As String, ByVal transferredBytes As Integer, ByVal totalBytes As Integer, ByVal message As String)
    Private Event onTransferProgress(ByVal source As String, ByVal destination As String, ByVal transferredBytes As Integer, ByVal totalBytes As Integer, ByVal message As String)
    Private Event onTransferEnd(ByVal source As String, ByVal destination As String, ByVal transferredBytes As Integer, ByVal totalBytes As Integer, ByVal message As String)

    Private Sub _onTransferStart(ByVal source As String, ByVal destination As String, ByVal transferredBytes As Integer, ByVal totalBytes As Integer, ByVal message As String) Handles transfer.OnTransferStart
        RaiseEvent onTransferStart(source, destination, transferredBytes, totalBytes, message)
    End Sub
    Private Sub _onTransferProgress(ByVal source As String, ByVal destination As String, ByVal transferredBytes As Integer, ByVal totalBytes As Integer, ByVal message As String) Handles transfer.OnTransferProgress
        RaiseEvent onTransferProgress(source, destination, transferredBytes, totalBytes, message)
    End Sub
    Private Sub _onTransferEnd(ByVal source As String, ByVal destination As String, ByVal transferredBytes As Integer, ByVal totalBytes As Integer, ByVal message As String) Handles transfer.OnTransferEnd
        RaiseEvent onTransferEnd(source, destination, transferredBytes, totalBytes, message)
    End Sub
End Class