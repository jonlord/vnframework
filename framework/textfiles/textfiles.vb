Imports System.IO
''' <summary>
''' Text File Manipulation Library for .NET
''' </summary>
Public Class textfiles
    ''' <summary>
    ''' Return all the contents of a file in a single string
    ''' </summary>
    ''' <param name="path">Path to the file</param>
    ''' <param name="errInfo">Holds error messages in caste they occur</param>
    ''' <returns>String representation of the contents of the file</returns>
    Public Shared Function getFileContents(ByVal path As String, Optional ByRef errInfo As String = "") As String
        Dim strContents As String
        Dim objReader As StreamReader
        Try
            objReader = New StreamReader(path)
            strContents = objReader.ReadToEnd()
            objReader.Close()
            Return strContents
        Catch Ex As Exception
            errInfo = Ex.Message
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Writes content into a file
    ''' </summary>
    ''' <param name="path">Path to the file</param>
    ''' <param name="contents">Contents to write</param>
    ''' <param name="errInfo">Holds error messages in caste they occur</param>
    Public Shared Sub writeFileContents(ByVal path As String, contents As String, Optional ByRef errInfo As String = "")
        Dim sw As New StreamWriter(path)
        sw.Write(contents)
        sw.Close()
    End Sub

    ''' <summary>
    ''' Removes the first line of a file
    ''' </summary>
    ''' <param name="path">Path to the file</param>
    Public Shared Sub removeFirstLine(path As String)
        Dim tempFileName As String = IO.Path.GetTempFileName() & ".txt"
        Dim sr As New StreamReader(path)
        Dim sw As New StreamWriter(tempFileName)

        For n As Int32 = 0 To 2
            Select Case n
                Case 0
                    sr.ReadLine()
                Case 1
                    sw.WriteLine(sr.ReadLine)
                Case Else
                    sw.Write(sr.ReadToEnd)
            End Select
        Next
        sw.Close()
        sr.Close()

        File.Copy(tempFileName, path, True)
        File.Delete(tempFileName)
    End Sub

    ''' <summary>
    ''' Changes the line feeds on a file to new lines
    ''' </summary>
    ''' <param name="path">Path to the file</param>
    ''' <param name="errInfo">Holds error messages in caste they occur</param>
    Public Shared Sub lineFeedsToNewLines(path As String, Optional ByRef errInfo As String = "")
        Dim sContents As String = getFileContents(path, errInfo)
        sContents = sContents.Replace(vbLf, vbNewLine)
        writeFileContents(path, sContents, errInfo)
    End Sub
End Class