Imports System.Windows.Forms
Imports System.Data.Odbc
Imports vnframework

Public Class CheckedListBoxEx
    Inherits CheckedListBox

    Private _sql As String

    Property selectFirst As Boolean = True
    Property cnx As OdbcConnection

    Private Sub populatFromSQL(ByVal query As String, Optional selectfirst As Boolean = True, Optional ByVal cnx As OdbcConnection = Nothing)
        With Me
            .Items.Clear()
            Try
                Dim rs As OdbcDataReader = ExecuteReader(query, cnx)
                While rs.Read
                    .Items.Add(New listItem(sql.Nz(rs(0), ""), sql.Nz(rs(1), "")))
                End While
                rs.Close()
            Catch thisError As OdbcException
                showError(thisError.Message)
            End Try
            If .Items.Count > 0 And selectfirst Then .SelectedIndex = 0
        End With
    End Sub

    Property fillQuery As String
        Get
            Return _sql
        End Get
        Set(value As String)
            If Not value Is Nothing Then
                If value.Length > 0 Then
                    If Not _cnx Is Nothing Then
                        populatFromSQL(value, selectFirst, _cnx)
                    Else
                        populatFromSQL(value, selectFirst)
                    End If
                End If
            End If
            _sql = value
        End Set
    End Property

    Property selectedId As String
        Get
            Try
                Return CType(SelectedItem, listItem).id
            Catch
                Return Nothing
            End Try
        End Get
        Set(value As String)
            Dim i As Integer = 0
            For Each item As listItem In Items
                If item.id = value Then
                    SelectedIndex = i
                    Exit For
                End If
                i += 1
            Next
        End Set
    End Property
    ReadOnly Property selectedName As String
        Get
            Try
                Return CType(SelectedItem, listItem).Name
            Catch
                Return Nothing
            End Try
        End Get
    End Property

    Public Sub refreshData()
        fillQuery = _sql
    End Sub
End Class