Imports System.Net
''' <summary>
''' FTP Library
''' </summary>
Public Class ftpLibrary
    ''' <summary>
    ''' Constructor fro FTP Class
    ''' </summary>
    ''' <param name="ftpaddress">IP Address or FQDN for FTP Server</param>
    ''' <param name="username">Username for FTP Server</param>
    ''' <param name="password">Password for FTP Server</param>
    Sub New(ftpaddress As String, username As String, password As String)
        Me.ftpAddress = ftpaddress
        Me.userName = username
        Me.password = password
    End Sub

    ''' <summary>
    ''' IP Address or FQDN for FTP Server
    ''' </summary>
    ''' <returns>IP Address or FQDN for FTP Server</returns>
    Property ftpAddress As String
    ''' <summary>
    ''' Username for FTP Server
    ''' </summary>
    ''' <returns>Username for FTP Server</returns>
    Property userName As String
    ''' <summary>
    ''' Password for FTP Server
    ''' </summary>
    ''' <returns>Password for FTP Server</returns>
    Property password As String

    ''' <summary>
    ''' Get a file from a FTP server and return its contents
    ''' </summary>
    ''' <param name="fileName">Remote path of file to retreive</param>
    ''' <returns>Contents of file</returns>
    Function getFile(fileName As String) As String
        getFile = ""
        Dim fwr As FtpWebRequest = FtpWebRequest.Create("ftp://" & ftpAddress)
        fwr.Credentials = New NetworkCredential(userName, password)
        fwr.KeepAlive = True
        fwr.Method = WebRequestMethods.Ftp.ListDirectory
        fwr.Proxy = Nothing

        Dim sr As New IO.StreamReader(fwr.GetResponse().GetResponseStream())
        Dim lst = sr.ReadToEnd().Split(vbNewLine)
        For Each file As String In lst
            file = file.Trim() 'remove any whitespace
            If file = ".." OrElse file = "." Then Continue For
            If file = fileName Then
                Dim fwr2 As Net.FtpWebRequest = FtpWebRequest.Create("ftp://" & ftpAddress & "/" & file)
                fwr2.Credentials = fwr.Credentials
                fwr2.KeepAlive = True
                fwr2.Method = WebRequestMethods.Ftp.DownloadFile
                fwr2.Proxy = Nothing

                Dim fileSR As New IO.StreamReader(fwr2.GetResponse().GetResponseStream())
                Dim fileData As String = fileSR.ReadToEnd()
                getFile = IO.Path.GetTempFileName & ".txt"
                files.save(getFile, fileData)
                fileSR.Close()
                Exit For
            End If
        Next

        sr.Close()
    End Function
End Class