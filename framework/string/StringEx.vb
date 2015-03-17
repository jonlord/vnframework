Imports System.Text.RegularExpressions
Imports System.Runtime.CompilerServices
Imports System.Security.Cryptography
Imports System.Text

''' <summary>
''' Extension to the String class allowing for additional functionality
''' </summary>
Public Module StringEx

    ''' <summary>
    ''' Strips the string from non-ASCII characters
    ''' </summary>
    ''' <param name="str">String to use</param>
    ''' <returns>The same string with non-ASCII characters removed</returns>
    <Extension()>
    Public Function removeNonASCII(str As String) As String
        removeNonASCII = Regex.Replace(str, "[^\i0000-\u007F]", String.Empty)
    End Function

    ''' <summary>
    ''' Returns an Arraylist from the separation of the string
    ''' </summary>
    ''' <param name="str">String to use</param>
    ''' <param name="separator">Separator character</param>
    ''' <returns>Arraylist contaning the string separated into elements</returns>
    <Extension()>
    Public Function toArrayList(str As String, Optional separator As String = ",") As ArrayList
        toArrayList = New ArrayList
        toArrayList.AddRange(Split(str, separator))
    End Function

    ''' <summary>
    ''' Determines if the string contains at least on memeber of the items contained in the ArrayList
    ''' </summary>
    ''' <param name="str">String to use</param>
    ''' <param name="matches">ArrayList of possible matches</param>
    ''' <returns>True if at least on memeber of the arraylist is contained in the string</returns>
    <Extension()>
    Public Function containsFromArrayList(str As String, matches As ArrayList, Optional caseSensitive As Boolean = False) As Boolean
        containsFromArrayList = False
        Dim strCopy As String = str
        For Each match As String In matches
            If caseSensitive = False Then
                strCopy = strCopy.ToLower
                match = match.ToLower
            End If
            If strCopy.Contains(match) Then _
                Return True
        Next
    End Function

    ''' <summary>
    ''' Escape characters for use in SQL queries
    ''' </summary>
    ''' <param name="str"></param>
    ''' <returns>The string with characters escaped</returns>
    <Extension()>
    Public Function escapeCharacters(str As String) As String
        escapeCharacters = str.Replace("'", "''")
    End Function

    ''' <summary>
    ''' Gets the MD5 hash for a string with optional salt
    ''' </summary>
    ''' <param name="str">String to be MD5'ed</param>
    ''' <param name="salt">Optional Salt string</param>
    ''' <returns>MD5 hash</returns>
    <Extension()>
    Public Function toMD5(str As String, Optional salt As String = "") As String
        str = str & salt
        'Create an encoding object to ensure the encoding standard for the source text
        Dim md5 As MD5 = MD5.Create()
        Dim inputBytes() As Byte = Encoding.ASCII.GetBytes(str)
        Dim hash() As Byte = md5.ComputeHash(inputBytes)

        Dim sb As New StringBuilder()
        For i As Integer = 0 To hash.Length - 1
            sb.Append(hash(i).ToString("X2"))
        Next i
        Return sb.ToString()
    End Function
End Module