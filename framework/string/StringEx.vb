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
    ''' Repeat a character n times
    ''' </summary>
    ''' <param name="str"></param>
    ''' <returns>The string with characters escaped</returns>
    <Extension()>
    Public Function replicate(str As String, numberOfTimes As Integer) As String
        replicate = New String(CChar(str), numberOfTimes)
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

    ''' <summary>
    ''' Converts a string to title case
    ''' </summary>
    ''' <param name="str">String to convert</param>
    ''' <returns>The string applying title case</returns>
    Public Function toTitleCase(ByVal str As String) As String
        str = Trim(str).ToLower
        If str = "" Then Return str

        Dim fCapitalizeNextLetter As Boolean = True
        Dim newStr As String = ""
        Dim sChar

        For i As Integer = 1 To Len(str)
            sChar = Mid(str, i, 1)
            If fCapitalizeNextLetter Then
                If isLetter(sChar) Then
                    sChar = UCase(sChar)
                    fCapitalizeNextLetter = False
                Else
                    fCapitalizeNextLetter = True
                End If
            Else
                If Not isLetter(sChar) Then fCapitalizeNextLetter = True
            End If
            newStr = newStr & sChar
        Next
        Return newStr
    End Function
    ''' <summary>
    ''' Determines if a character is a letter
    ''' </summary>
    ''' <param name="sChar">Character to test</param>
    ''' <returns>True if it is a letter, false if it is not</returns>
    Function isLetter(ByVal sChar As String) As Boolean
        isLetter = False
        Dim nASCII As Integer = Asc(sChar)
        If ((nASCII >= 65) And (nASCII <= 90)) Then isLetter = True ' Upper case letters
        If ((nASCII >= 97) And (nASCII <= 122)) Then isLetter = True ' Lower case letters
        If nASCII = 241 Or nASCII = 225 Or nASCII = 233 Or nASCII = 237 Or nASCII = 250 Then isLetter = True 'ñ á é í ú
    End Function

End Module