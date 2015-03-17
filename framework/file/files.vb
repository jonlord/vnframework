Imports System.IO

''' <summary>
''' File helpder functions
''' </summary>
Public Class files
    ''' <summary>
    ''' Saves the desired contents to a text file
    ''' </summary>
    ''' <param name="fileName">File name with complete path</param>
    ''' <param name="contents">Contents to be saved</param>
    ''' <returns>True if the operation was successful or false if it failed</returns>
    Shared Function save(ByVal fileName As String, ByVal contents As String) As Boolean
        Try
            Dim sw As StreamWriter = File.CreateText(fileName)
            sw.Write(contents)
            sw.Close()
            sw.Dispose()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    ''' <summary>
    ''' Saves the ontents of a file to an arraylist one line per file
    ''' </summary>
    ''' <param name="fileName">File name with complete path</param>
    ''' <returns>Arraylist containing all the liens of a file</returns>
    Shared Function toArrayList(fileName As String) As ArrayList
        toArrayList = New ArrayList
        Dim sLine As String
        Try
            Dim objReader As New StreamReader(fileName)
            Do
                sLine = objReader.ReadLine()
                If Not sLine Is Nothing Then
                    toArrayList.Add(sLine)
                End If
            Loop Until sLine Is Nothing
            objReader.Close()
            objReader.Dispose()
        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try
    End Function
    ''' <summary>
    ''' Determine if a file exists
    ''' </summary>
    ''' <param name="fileName">File name with complete path</param>
    ''' <returns>True if it exists; false if not</returns>
    Public Shared Function exists(ByVal fileName As String) As Boolean
        If fileName.Length = 0 Then Return False
        Dim f As New FileInfo(fileName)
        Return f.Exists
    End Function
    ''' <summary>
    ''' Determine if a directory exists
    ''' </summary>
    ''' <param name="filename">File name with complete path</param>
    ''' <returns>True if it exists; false if not</returns>
    Shared Function directoryExists(ByVal filename As String) As Boolean
        If filename.Length = 0 Then Return False
        Return Directory.Exists(filename)
    End Function
    ''' <summary>
    ''' Gets the size of a file in bytes
    ''' </summary>
    ''' <param name="filename">File name with complete path</param>
    ''' <returns>Size in bytes</returns>
    Public Shared Function size(ByVal filename As String) As Long
        Try
            Dim MyFile As New FileInfo(filename)
            Dim FileSize As Long = MyFile.Length
            Return FileSize
        Catch ex As Exception
            Return 0
        End Try
    End Function
    ''' <summary>
    ''' Reads the file and returns a byte array with its content
    ''' </summary>
    ''' <param name="fileName">File name with complete path</param>
    ''' <returns>Byte array with the file's content</returns>
    Shared Function read(ByVal fileName As String) As Byte()
        Dim fs As FileStream = New FileStream(fileName, FileMode.Open, FileAccess.Read)
        Dim FileSize As Integer = fs.Length
        Dim fileReader As Byte() = New Byte(FileSize) {}
        fs.Read(fileReader, 0, FileSize)
        fs.Close()

        Return fileReader
    End Function

    ''' <summary>
    ''' Sorts a FileInfo array by creation date
    ''' </summary>
    ''' <param name="fileList">FileInfo Array to sort</param>
    Public Shared Sub sortFileListByDate(ByRef fileList() As FileInfo)
        Dim comparer As New dateComparer
        Array.Sort(fileList, comparer)
    End Sub

    ''' <summary>
    ''' Determines if file1 is newer than file2
    ''' </summary>
    ''' <param name="file1">First file to compare</param>
    ''' <param name="file2">Second file to compare</param>
    ''' <returns>True if file1 is newer, false if older o same age</returns>
    Shared Function isNewer(file1 As FileInfo, file2 As FileInfo) As Boolean
        Dim d As New dateComparer
        Dim result As Integer = d.Compare(file1, file2)
        If result = 1 Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Class dateComparer
        Implements IComparer

        Public Function Compare(ByVal info1 As Object, ByVal info2 As Object) As Integer Implements IComparer.Compare
            Dim FileInfo1 As FileInfo = DirectCast(info1, FileInfo)
            Dim FileInfo2 As FileInfo = DirectCast(info2, FileInfo)

            Dim Date1 As DateTime = FileInfo1.CreationTime
            Dim Date2 As DateTime = FileInfo2.CreationTime

            If Date1 > Date2 Then Return 1
            If Date1 < Date2 Then Return -1
            Return 0
        End Function
    End Class
End Class