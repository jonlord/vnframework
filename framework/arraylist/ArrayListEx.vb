Imports System.IO

''' <summary>
''' ArrayList Extension
''' </summary>
Public Class ArrayListEx
    Inherits ArrayList
    ''' <summary>
    ''' Joins all the elements of an ArrayList as a single string separated by glue
    ''' </summary>
    ''' <param name="glue">Separator between elements of the ArrayList</param>
    ''' <returns>String representation of the elements of the ArrayList</returns>
    Function implode(Optional glue As String = ",") As String
        implode = ""
        For Each a As String In Me
            implode &= a & glue
        Next
        If implode.EndsWith(glue) Then implode = implode.Substring(0, implode.Length - glue.Length)
    End Function

    ''' <summary>
    ''' Sorts and removes duplicate values
    ''' </summary>
    Sub removeDuplicates()
        With Me
            .Sort()
            Dim count As Long = .Count
            For i As Integer = count - 1 To 1 Step -1
                If .Item(i).ToString() = .Item(i - 1).ToString() Then
                    .RemoveAt(i)
                End If
            Next i
        End With
    End Sub

    Sub toTextFile(fileName As String)
        Dim theWriter As New StreamWriter(fileName)
        For Each currentElement As String In Me
            theWriter.WriteLine(currentElement)
        Next
        theWriter.Close()
    End Sub

    ''' <summary>
    ''' Initializes a new instance of the ArrayListEx class that contains elements copied from the specified collection and that has the same initial capacity as the number of elements copied.
    ''' </summary>
    ''' <param name="c">The System.Collections.ICollection whose elements are copied to the new list.</param>
    Public Sub New(c As ICollection)
        MyBase.New(c)
    End Sub

    ''' <summary>
    ''' Initializes a new instance of the ArrayListEx class that is empty and has the specified initial capacity.
    ''' </summary>
    ''' <param name="capacity">The number of elements that the new list can initially store.</param>
    Public Sub New(capacity As Integer)
        MyBase.New(capacity)
    End Sub
    ''' <summary>
    ''' Initializes a new instance of the ArrayListEx class that is empty
    ''' </summary>
    Sub New()
        MyBase.New
    End Sub
End Class