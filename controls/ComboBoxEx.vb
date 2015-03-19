Imports System.Windows.Forms
Imports System.Data.Odbc
Imports vnframework

Public Class ComboBoxEx
    Inherits ComboBox

    Private _sql As String
    Private _cnx As OdbcConnection

    Property selectFirst As Boolean = True

    Property connection As OdbcConnection
        Get
            Return _cnx
        End Get
        Set(value As OdbcConnection)
            _cnx = value
        End Set
    End Property

    Property fillQuery As String
        Get
            Return _sql
        End Get
        Set(value As String)
            If Not value Is Nothing Then
                If connection Is Nothing Then _
                    connection = getGlobalConnection()
                If value.Length > 0 Then _
                        populate(value, connection, selectFirst)
            End If
            _sql = value
        End Set
    End Property

    Property SelectedId As String
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
    ReadOnly Property SelectedName As String
        Get
            Try
                Return CType(SelectedItem, listItem).Name
            Catch
                Return Nothing
            End Try
        End Get
    End Property

    Public Property AllowNewInput As Boolean = False

    Public Shared Sub cbo_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyUp
        Dim cbo As ComboBox = DirectCast(sender, ComboBox)
        Select Case e.KeyCode
            Case Keys.Back, Keys.Left, Keys.Right, Keys.Up, Keys.Delete, Keys.Down, Keys.Tab
                Return
        End Select
        Dim ComboBoxText As String = Strings.Left(cbo.Text, cbo.SelectionStart)
        Dim FirstFound As Integer = cbo.FindString(ComboBoxText)
        If FirstFound >= 0 Then
            Dim FirstFoundItem As Object = cbo.Items(FirstFound)
            Dim FirstFoundText As String = cbo.GetItemText(FirstFoundItem)
            cbo.Text = FirstFoundText
            cbo.SelectionStart = ComboBoxText.Length
            cbo.SelectionLength = cbo.Text.Length
        End If
    End Sub
    Public Sub cbo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Leave
        Dim cbo As ComboBox = DirectCast(sender, ComboBox)
        If cbo.SelectedIndex = -1 And Not AllowNewInput Then cbo.Text = ""
    End Sub

    Public Sub refreshData()
        fillQuery = _sql
    End Sub

    Private Sub populate(ByVal query As String, ByVal connection As OdbcConnection, Optional selectfirst As Boolean = True)
        With Me
            .Items.Clear()
            Try
                Dim rs As OdbcDataReader = ExecuteReader(query, connection)
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
End Class