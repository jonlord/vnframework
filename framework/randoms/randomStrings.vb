''' <summary>
''' Random string manipulation library
''' </summary>
Public Class randomStrings
    ''' <summary>
    ''' Generates a list of random string of a desired length
    ''' </summary>
    ''' <param name="amount">Amount of elements to generate</param>
    ''' <param name="length">Length of each element</param>
    ''' <param name="allowDuplicates">Allow duplicate items</param>
    ''' <returns>ArrayList consisting of random strings</returns>
    Shared Function toArrayList(amount As Long, length As Integer, Optional allowDuplicates As Boolean = False) As ArrayList
        Const letters As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        Dim lettersLength As Integer = letters.Length
        Dim r As New System.Random

        toArrayList = New ArrayList
        With toArrayList
            Dim pendingCount As Integer = Math.Max(amount - .Count, 0)
            While pendingCount > 0
                For i As Long = 1 To pendingCount
                    Dim str As String = ""
                    For j As Integer = 1 To length
                        str &= letters.Substring(r.Next(0, lettersLength), 1)
                    Next j
                    .Add(str)
                Next i
                .Sort()

                If Not allowDuplicates Then
                    Dim count As Long = .Count
                    For i As Integer = count - 1 To 1 Step -1
                        If .Item(i) = .Item(i - 1) Then
                            .RemoveAt(i)
                        End If
                    Next i
                End If
                pendingCount = Math.Max(amount - .Count, 0)
            End While
        End With
    End Function
End Class