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
End Class
