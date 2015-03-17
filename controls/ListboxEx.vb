Imports System.Windows.Forms
Imports System.Data.Odbc
Imports vnframework

Public Class ListBoxEx
    Inherits ListBox

    Private _fillSQL As String
    Private _deleteSQL As String
    Private _saveSQL As String
    Private originalItems As New ArrayList
    Property selectFirst As Boolean = True
    Property cnx As OdbcConnection

    Property fillQuery As String
        Get
            Return _fillSQL
        End Get
        Set(value As String)
            If Not value Is Nothing Then
                If value.Length > 0 Then
                    _fillSQL = value
                    populate()
                End If
            End If
        End Set
    End Property
    Property saveQuery As String
        Get
            Return _saveSQL
        End Get
        Set(value As String)
            _saveSQL = value
        End Set
    End Property
    Property deleteQuery As String
        Get
            Return _deleteSQL
        End Get
        Set(value As String)
            _deleteSQL = value
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

    ReadOnly Property toCSV As String
        Get
            Dim list As New ArrayListEx
            For Each item As listItem In SelectedItems
                list.Add(item.id)
            Next
            Return list.implode()
        End Get
    End Property

    Public Sub refreshData()
        fillQuery = _fillSQL
    End Sub

    Public Shared Sub popular(ByVal lb As ListBoxEx, ByVal query As String, Optional selectfirst As Boolean = True, Optional ByVal cnx As OdbcConnection = Nothing)
        lb.Items.Clear()
        Try
            Dim rs As OdbcDataReader = ExecuteReader(query, cnx)
            While rs.Read
                lb.Items.Add(New listItem(sql.Nz(rs(0), ""), sql.Nz(rs(1), "")))
            End While
            rs.Close()
        Catch thisError As OdbcException
            showError(thisError.Message)
        End Try
        If lb.Items.Count > 0 And selectfirst Then lb.SelectedIndex = 0
    End Sub
    Public Sub populate()
        With Me
            originalItems.Clear()
            .Items.Clear()
            Try
                Dim rs As OdbcDataReader = ExecuteReader(_fillSQL, cnx)
                While rs.Read
                    Dim newItem As listItem = New listItem(sql.Nz(rs(0), ""), sql.Nz(rs(1), ""))
                    .Items.Add(newItem)
                    originalItems.Add(newItem)
                End While
                rs.Close()
            Catch thisError As OdbcException
                showError(thisError.Message)
            End Try
            If .Items.Count > 0 And selectFirst Then .SelectedIndex = 0
        End With
    End Sub
    Public Shared Function valor(ByVal lb As Object) As Integer
        With CType(lb.selecteditem, listItem)
            Return .id
        End With
    End Function
    Public Shared Sub seleccionar(ByVal lb As Object, ByVal value As Object)
        lb.SelectedIndex = -1
        For Each item In lb.Items
            With CType(item, listItem)
                If .id = value Then
                    lb.SelectedItem = item
                    Exit Sub
                End If
            End With
        Next
    End Sub

    Public Function saveSQL() As ArrayList
        saveSQL = New ArrayList
        For Each item As listItem In Me.Items
            If originalItems.Contains(item) Then
                originalItems.Remove(item)
            Else
                saveSQL.Add(saveQuery.Replace("@@id@@", item.id))
            End If
        Next
        For Each item As listItem In originalItems
            saveSQL.Add(deleteQuery.Replace("@@id@@", item.id))
        Next
    End Function
End Class