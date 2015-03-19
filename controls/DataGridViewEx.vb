Imports System.Data.Odbc
Imports System.Data.SqlClient
Imports MySql.Data.MySqlClient
Imports System.Windows.Forms

Public Class DatagridviewEx
    Inherits DataGridView

    ''' <summary>
    ''' Formats a DataGridView
    ''' </summary>
    Sub format()
        Dim myBackColor As Drawing.Color = Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        With Me
            Dim DataGridViewCellStyle1 As DataGridViewCellStyle = New DataGridViewCellStyle()
            DataGridViewCellStyle1.BackColor = myBackColor
            .AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        End With
    End Sub

    ''' <summary>
    ''' Populates a DataGridView with a DataTable
    ''' </summary>
    ''' <param name="dataTable">Datatable to use</param>
    ''' <param name="columnNames">List of Column Names</param>
    Sub populate(dataTable As DataTable, ByVal columnNames As String())
        Try
            With Me
                .Rows.Clear()
                .Columns.Clear()

                For i As Integer = 0 To columnNames.Count - 1
                    If dataTable.Columns(i).DataType = GetType(Boolean) Then
                        Dim col As New DataGridViewCheckBoxColumn()
                        col.Name = dataTable.Columns(i).ColumnName
                        col.HeaderText = columnNames(i)
                        .Columns.Add(col)
                    Else
                        .Columns.Add(dataTable.Columns(i).ColumnName, columnNames(i))
                    End If
                Next

                For Each dataRow1 As DataRow In dataTable.Rows
                    Dim dataGridRow1 As New DataGridViewRow
                    dataGridRow1.CreateCells(Me)
                    For i As Integer = 0 To columnNames.Count - 1
                        dataGridRow1.Cells(i).Value = dataRow1(i)
                    Next
                    .Rows.Add(dataGridRow1)
                Next

                For i As Integer = 0 To .Columns.Count - 1
                    If Not columnNames Is Nothing Then
                        If columnNames.Count - 1 >= i Then
                            .Columns(i).ReadOnly = True
                            .Columns(i).HeaderCell = Nothing
                            .Columns(i).HeaderCell = New DataGridViewAutoFilterColumnHeaderCell()
                            .Columns(i).HeaderText = columnNames(i)
                        End If
                    End If
                Next i
                .AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)
                .ReadOnly = True
                .AllowUserToAddRows = False
                .AllowUserToDeleteRows = False
            End With
        Catch ex As Exception
            showError(ERRORLOADINGDATA & vbNewLine & vbNewLine & ex.Message)
            logs.writeToErrorLog(ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Populate a DataGridView from MySQL
    ''' </summary>
    ''' <param name="query">Select query to use</param>
    ''' <param name="columnNames">Column Names</param>
    ''' <param name="bsBindignSource">A BindingSource </param>
    ''' <param name="connection">a MySQL Connection</param>
    Sub populateMySQL(ByVal query As String, ByVal columnNames As String(), ByRef bsBindignSource As BindingSource, ByVal connection As MySqlConnection)
        Dim dataAdapter As MySqlDataAdapter

        If connection.State <> ConnectionState.Open Then connection.Open()
        dataAdapter = New MySqlDataAdapter(query, connection)
        Try
            Dim table As DataTable = New DataTable() With {.Locale = System.Globalization.CultureInfo.InvariantCulture}
            dataAdapter.Fill(table)
            bsBindignSource.DataSource = table
            table.Dispose()
            With Me
                .AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)
                .DataSource = bsBindignSource

                For i As Integer = 0 To columnNames.Count - 1
                    If Not columnNames Is Nothing Then
                        If columnNames.Count - 1 >= i Then
                            .Columns(i).ReadOnly = True
                            .Columns(i).HeaderCell = Nothing
                            .Columns(i).HeaderCell = New DataGridViewAutoFilterColumnHeaderCell()
                            .Columns(i).HeaderText = columnNames(i)
                        End If
                    End If
                Next i
                .AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)
                .ReadOnly = True
                .AllowUserToAddRows = False
                .AllowUserToDeleteRows = False
            End With
            connection.Close()
        Catch ex As MySqlException
            showError(ERRORLOADINGDATA & vbNewLine & vbNewLine & ex.Message)
            logs.writeToLog(ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Populate a DataGridView from a SO
    ''' </summary>
    ''' <param name="query">Select query to use</param>
    ''' <param name="columnNames">Column Names</param>
    ''' <param name="bsBindignSource">A BindingSource </param>
    ''' <param name="connection">a SQL Connection</param>
    Sub populateSP(ByVal query As String, ByVal columnNames As String(), ByRef bsBindignSource As BindingSource, connection As SqlConnection)
        Try
            Dim selectCMD As New SqlCommand(query, connection)
            selectCMD.CommandTimeout = 60 * 10
            Using dataAdapter As SqlDataAdapter = New SqlDataAdapter(selectCMD)
                Dim table As DataTable = New DataTable() With {.Locale = Globalization.CultureInfo.InvariantCulture}
                dataAdapter.Fill(table)
                bsBindignSource.DataSource = table
                table.Dispose()
            End Using
            With Me
                .AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)
                .DataSource = bsBindignSource

                For i As Integer = 0 To .Columns.Count - 1
                    If columnNames.Count >= i + 1 Then
                        .Columns(i).HeaderText = columnNames(i)
                        .Columns(i).ReadOnly = True
                    End If
                Next i
                .AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)
                .ReadOnly = True
                .AllowUserToAddRows = False
                .AllowUserToDeleteRows = False
            End With
            selectCMD.Dispose()
        Catch ex As SqlException
            showError(ERRORLOADINGDATA & vbNewLine & vbNewLine & ex.Message)
            logs.writeToLog(ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Populate a DataGridView
    ''' </summary>
    ''' <param name="query">Select query to use</param>
    ''' <param name="columnNames">Column Names</param>
    ''' <param name="bsBindignSource">A BindingSource </param>
    Sub populate(ByVal query As String, ByVal columnNames As String(), Optional ByRef bsBindignSource As BindingSource = Nothing)
        populate(query, columnNames, bsBindignSource, Nothing)
    End Sub

    ''' <summary>
    ''' Populate a DataGridView from ODBC
    ''' </summary>
    ''' <param name="query">Select query to use</param>
    ''' <param name="columnNames">Column Names</param>
    ''' <param name="bsBindignSource">A BindingSource </param>
    ''' <param name="connection">a ODBC Connection</param>
    Sub populate(ByVal query As String, ByVal columnNames As String(), ByRef bsBindignSource As BindingSource, connection As OdbcConnection)
        If connection Is Nothing Then connection = getGlobalConnection()
        If bsBindignSource Is Nothing Then bsBindignSource = New BindingSource
        Try
            query = sql.parseToSQLServer(query)
            If query.StartsWith("EXEC") Then
                populateSP(query, columnNames, bsBindignSource, getGlobalSQLConnection())
                Exit Sub
            End If

            Dim selectCMD As New OdbcCommand(query, connection)
            selectCMD.CommandTimeout = 240 * 1000
            Me.Columns.Clear()
            bsBindignSource.Filter = ""
            'dbGridView.Rows.Clear()
            Using dataAdapter As OdbcDataAdapter = New OdbcDataAdapter(selectCMD)
                Dim table As DataTable = New DataTable() With {.Locale = System.Globalization.CultureInfo.InvariantCulture}
                dataAdapter.Fill(table)
                bsBindignSource.DataSource = table
                table.Dispose()
            End Using
            With Me
                .AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)
                .DataSource = bsBindignSource

                For i As Integer = 0 To .Columns.Count - 1
                    If Not columnNames Is Nothing Then
                        If columnNames.Count - 1 >= i Then
                            .Columns(i).ReadOnly = True
                            .Columns(i).HeaderCell = Nothing
                            .Columns(i).HeaderCell = New DataGridViewAutoFilterColumnHeaderCell()
                            .Columns(i).HeaderText = columnNames(i)
                        End If
                    End If
                Next i
                .AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)
                .ReadOnly = True
                .AllowUserToAddRows = False
                .AllowUserToDeleteRows = False
            End With
            selectCMD.Dispose()
        Catch ex As OdbcException
            showError(ERRORLOADINGDATA & vbNewLine & vbNewLine & ex.Message)
            logs.writeToLog(ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Populate a DataGridView from CSV
    ''' </summary>
    ''' <param name="filePath">Path to file</param>
    ''' <param name="columnNames">Column Names</param>
    ''' <param name="conditions">Conditinos to filter out records</param>
    Sub populateFromCSV(ByVal filePath As String, ByVal columnNames As String(), conditions As String())
        Try
            With Me
                .Rows.Clear()
                .Columns.Clear()

                Dim textLine As String, splitLine() As String

                For i As Integer = 0 To columnNames.Count - 1
                    .Columns.Add("col" & i, columnNames(i))
                Next

                If IO.File.Exists(filePath) = True Then
                    Dim objReader As New IO.StreamReader(filePath)
                    Do While objReader.Peek() <> -1
                        textLine = objReader.ReadLine()
                        splitLine = Split(textLine, ",")

                        Dim add As Boolean = IIf(conditions.Count = 0, True, False)
                        For i As Integer = 0 To conditions.Count - 1
                            If conditions.Count >= i Then If conditions(i) <> "" Then If splitLine(i) = conditions(i) Then add = True
                        Next
                        If add Then .Rows.Add(splitLine)
                    Loop
                    objReader.Dispose()
                End If

                '    .AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)
                '    .DataSource = bsBindignSource

                '    For i As Integer = 0 To .Columns.Count - 1
                '        If columnNames.Count - 1 >= i Then
                '            .Columns(i).HeaderText = columnNames(i)
                '            .Columns(i).ReadOnly = True
                '        End If
                '    Next i
                .AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)
                .ReadOnly = True
                .AllowUserToAddRows = False
                .AllowUserToDeleteRows = False
            End With
        Catch ex As Exception
            showError(ERRORLOADINGDATA & vbNewLine & vbNewLine & ex.Message)
            logs.writeToLog(ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Change a DataGridView's column to a combobox
    ''' </summary>
    ''' <param name="SQL">Query to use for data</param>
    ''' <param name="boundColumn">DataGridView column to replace</param>
    ''' <param name="displayMember">Display member of combobox</param>
    ''' <param name="valueMember">Value memeber of combobox</param>
    Sub columnToCombobox(SQL As String, boundColumn As String, valueMember As String, displayMember As String)

        Dim dt As New DataTable()
        dt.Load(ExecuteReader(String.Format(SQL), getGlobalConnection))
        Dim headerText As String = Me.Columns(boundColumn).HeaderText
        Me.Columns(boundColumn).Visible = False
        Dim column As New DataGridViewComboBoxColumn
        With column
            .DataSource = dt
            .DisplayMember = displayMember
            .ValueMember = valueMember
            .HeaderText = headerText
            .Name = boundColumn
            .DataPropertyName = boundColumn
        End With
        Dim indice As Integer = 2
        Me.Columns.Insert(indice, column)
    End Sub
End Class