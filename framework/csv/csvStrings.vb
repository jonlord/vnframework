Imports System.Text

''' <summary>
''' CSV manipulation Library for .NET
''' </summary>
Public Class csvStrings
    ''' <summary>
    ''' String builder used to manage the csv to be returned
    ''' </summary>
    Private CSVBuilder As StringBuilder

    ''' <summary>
    ''' Identifies if column headers are used
    ''' </summary>
    ''' <returns>Boolean that determines if column headers are used</returns>
    Public Property useColumnHeaders As Boolean = True
    ''' <summary>
    ''' Text Qualifier
    ''' </summary>
    ''' <returns>Character used as text qualifier</returns>
    Public Property textQualifiers As Char = """"c
    ''' <summary>
    ''' Field delimiter
    ''' </summary>
    ''' <returns>Character used as a text delimiter between fields</returns>
    Public Property textDelimiter As Char = ","c

    ''' <summary>
    ''' Converts inputTable into a CSV string
    ''' </summary>
    ''' <param name="inputTable">Data Table to convert</param>
    ''' <returns>datatable as a CSV string</returns>
    Public Function fromDataTable(ByVal inputTable As DataTable) As String
        CSVBuilder = New StringBuilder()
        If useColumnHeaders Then
            createHeader(inputTable)
        End If
        createRows(inputTable)
        Return CSVBuilder.ToString()
    End Function
    ''' <summary>
    ''' Add the contents to the CSV string
    ''' </summary>
    ''' <param name="inputTable">Data table</param>
    Private Sub createRows(ByVal inputTable As DataTable)
        For Each ExportRow As DataRow In inputTable.Rows
            For Each ExportColumn As DataColumn In inputTable.Columns
                Dim ColumnText As String = ExportRow(ExportColumn.ColumnName).ToString()
                ColumnText = ColumnText.Replace(textQualifiers.ToString(), textQualifiers.ToString() + textQualifiers.ToString())
                CSVBuilder.Append(textQualifiers + ColumnText + textQualifiers)
                CSVBuilder.Append(textDelimiter)
            Next
            CSVBuilder.AppendLine()
        Next
    End Sub
    ''' <summary>
    ''' Adds the header to the CSV string
    ''' </summary>
    ''' <param name="inputTable">Data Table</param>
    Private Sub createHeader(ByVal inputTable As DataTable)
        For Each ExportColumn As DataColumn In inputTable.Columns
            Dim ColumnText As String = ExportColumn.ColumnName.ToString()
            ColumnText = ColumnText.Replace(textQualifiers.ToString(), textQualifiers.ToString() + textQualifiers.ToString())
            CSVBuilder.Append(textQualifiers + ExportColumn.ColumnName + textQualifiers)
            CSVBuilder.Append(textDelimiter)
        Next
        CSVBuilder.AppendLine()
    End Sub
End Class